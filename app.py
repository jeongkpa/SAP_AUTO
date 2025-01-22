import streamlit as st
import pyautogui
import cv2
import numpy as np
from screeninfo import get_monitors
import os
from datetime import datetime
import win32gui
import win32ui
import win32con
from ctypes import windll

def ensure_temp_dir():
    """임시 디렉토리 생성"""
    temp_dir = "./temp"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    return temp_dir

def capture_screen(region):
    """
    win32gui를 사용하여 화면을 캡처합니다.
    region: (x, y, width, height) 튜플
    """
    # 캡처할 영역의 좌표와 크기
    x, y, width, height = region
    
    # 화면 DC 가져오기
    hwnd = win32gui.GetDesktopWindow()
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    
    # 비트맵 생성
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)
    
    # 화면 복사
    result = saveDC.BitBlt((0, 0), (width, height), mfcDC, (x, y), win32con.SRCCOPY)
    if result == None:
        print("BitBlt failed!")
    
    # 비트맵을 numpy 배열로 변환
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    
    img = np.frombuffer(bmpstr, dtype='uint8')
    img.shape = (height, width, 4)  # RGBA
    
    # 리소스 해제
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    
    # RGBA to BGR 변환
    img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    return img

def run_macro(selected_monitor_index, working_text, warning_text, status_container):
    """매크로 실행 함수"""
    monitors = list(get_monitors())
    selected_monitor = monitors[selected_monitor_index]
    
    # 모니터 정보 출력
    st.write("=== 모니터 정보 ===")
    for i, m in enumerate(monitors):
        st.write(f"모니터 {i+1}: x={m.x}, y={m.y}, width={m.width}, height={m.height}")
    
    # 캡처 영역 설정
    region = (
        selected_monitor.x,
        selected_monitor.y,
        selected_monitor.width,
        selected_monitor.height
    )
    
    # 임시 디렉토리 생성
    temp_dir = ensure_temp_dir()
    
    try:
        # 스크린샷 캡처
        screenshot_cv = capture_screen(region)
        
        # 템플릿 매칭 및 처리
        template = cv2.imread("./asset/box.png", cv2.IMREAD_COLOR)
        if template is None:
            st.error("box.png 파일을 찾을 수 없거나 이미지가 아닙니다.")
            return
        
        # 템플릿 이미지 크기 가져오기
        th, tw = template.shape[:2]
        print(f"템플릿 이미지 크기: {tw}x{th}")
            
        # 스크롤 처리
        scroll_count = 0
        items_found = True
        total_matches = 0  # 총 매칭 수 초기화
        consecutive_no_matches = 0
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        while scroll_count < st.session_state.max_scrolls and items_found:
            try:
                # 현재 화면 캡처
                screenshot_cv = capture_screen(region)
                
                # 스크린샷 저장 (타임스탬프에 스크롤 횟수 추가)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_path = os.path.join(temp_dir, f"screenshot_{timestamp}_scroll{scroll_count}.png")
                cv2.imwrite(screenshot_path, screenshot_cv)
                print(f"스크린샷 저장됨: {screenshot_path}")
                
            except Exception as e:
                print(f"스크린샷 캡처 중 오류 발생: {str(e)}")
                st.error(f"스크린샷 캡처 실패: {str(e)}")
                return

            # 템플릿 매칭
            result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
            threshold = st.session_state.threshold
            loc = np.where(result >= threshold)
            
            # 현재 화면에서 매칭된 위치 표시
            debug_image = screenshot_cv.copy()
            match_count = 0
            
            # 현재 화면에서 찾은 체크박스 클릭
            for pt_y, pt_x in zip(*loc):
                match_count += 1
                total_matches += 1  # 전체 매칭 수 증가
                # 매칭된 영역 표시
                cv2.rectangle(debug_image, (pt_x, pt_y), (pt_x + tw, pt_y + th), (0, 0, 255), 2)
                
                # 클릭 위치 계산 및 클릭
                center_x = region[0] + pt_x + tw // 2
                center_y = region[1] + pt_y + th // 2
                print(f"클릭 위치: ({center_x}, {center_y})")
                pyautogui.click(center_x, center_y)
                
                # 클릭 후 잠시 대기 (1/3로 줄임)
                pyautogui.sleep(0.17)  # 0.5 -> 0.17

            # 디버그 이미지 저장
            debug_path = os.path.join(temp_dir, f"debug_visualization_{timestamp}_scroll{scroll_count}.png")
            cv2.imwrite(debug_path, debug_image)
            print(f"디버그 이미지 저장됨: {debug_path}")
            print(f"현재 화면에서 매칭된 항목 수: {match_count}")

            # 현재 화면에서 항목을 찾지 못했다면 스크롤
            if match_count == 0:
                consecutive_no_matches += 1
                if consecutive_no_matches >= 4:  # 4번 연속으로 매칭 실패
                    working_text.markdown('<p class="working-text">✅ 작업 완료!</p>', unsafe_allow_html=True)
                    warning_text.empty()  # 경고 메시지 제거
                    status_container.warning("3번 연속으로 매칭된 항목을 찾지 못해 중단합니다.")
                    break
                    
                if scroll_count >= st.session_state.max_scrolls - 1:
                    items_found = False
                    break
                
                status_text.write(f"스크롤 {scroll_count + 1}/{st.session_state.max_scrolls} (연속 미발견: {consecutive_no_matches})")
                pyautogui.scroll(st.session_state.scroll_amount)
                scroll_count += 1
                pyautogui.sleep(1)
            else:
                consecutive_no_matches = 0  # 매칭 성공시 카운터 리셋
                print(f"현재 화면에서 {match_count}개 항목 처리 완료")

            # 매칭된 결과 로그만 출력
            if match_count > 0:
                print(f"스크롤 {scroll_count + 1}: {match_count}개 항목 발견")
            
            status_text.write(f"진행 중... (스크롤: {scroll_count + 1}/{st.session_state.max_scrolls})")
            progress_bar.progress((scroll_count + 1) / st.session_state.max_scrolls)

        # 작업 완료 시 UI 업데이트
        working_text.markdown('<p class="working-text">✅ 작업 완료!</p>', unsafe_allow_html=True)
        warning_text.empty()  # 경고 메시지 제거
        
        if consecutive_no_matches >= 3:
            status_container.warning("3번 연속으로 매칭된 항목을 찾지 못해 중단합니다.")
            st.info(f"작업이 중단되었습니다. (3번 연속 미발견)\n총 {total_matches}개의 체크박스를 작업했습니다.")
        else:
            st.success(f"작업이 완료되었습니다!\n총 {total_matches}개의 체크박스를 작업했습니다.")
        
    except Exception as e:
        st.error(f"오류 발생: {str(e)}")

def show_working_animation():
    """작업 중 애니메이션 표시"""
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.image("asset/working.gif", use_column_width=True)

# Streamlit UI
def main():
    st.title("입고 구매 오더 자동화")
    
    # CSS 스타일 추가
    st.markdown("""
        <style>
        .warning-text {
            color: red;
            font-size: 24px;
            font-weight: bold;
            text-align: center;
            padding: 20px;
        }
        .working-text {
            color: #1E88E5;
            font-size: 28px;
            font-weight: bold;
            text-align: center;
            padding: 20px;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # session_state 초기화
    if 'max_scrolls' not in st.session_state:
        st.session_state.max_scrolls = 5
    if 'scroll_amount' not in st.session_state:
        st.session_state.scroll_amount = -200
    if 'threshold' not in st.session_state:
        st.session_state.threshold = 0.95
    
    # 사이드바에 설정 옵션들 추가
    with st.sidebar:
        st.header("설정")
        
        # 모니터 선택
        monitors = list(get_monitors())
        monitor_options = [f"모니터 {i+1} ({m.width}x{m.height})" for i, m in enumerate(monitors)]
        selected_monitor = st.selectbox(
            "모니터 선택",
            range(len(monitor_options)),
            format_func=lambda x: monitor_options[x]
        )
        
        # 매크로 설정
        st.session_state.max_scrolls = st.slider("최대 스크롤 횟수", 1, 20, st.session_state.max_scrolls)
        st.session_state.scroll_amount = st.slider("스크롤 크기", -500, -100, st.session_state.scroll_amount)
        st.session_state.threshold = st.slider("매칭 임계값", 0.0, 1.0, st.session_state.threshold)
        
        # 작업 시작 버튼
        start_button = st.button("작업 시작", use_container_width=True)
    
    # 메인 화면에 진행 상태 표시
    main_container = st.container()
    
    if start_button:
        with main_container:
            # 작업 중 텍스트와 경고 메시지를 placeholder로 설정
            working_text = st.empty()
            warning_text = st.empty()
            status_container = st.empty()
            
            # 초기 메시지 표시
            working_text.markdown('<p class="working-text">🔄 작업 중...</p>', unsafe_allow_html=True)
            warning_text.markdown("""
                <div class="warning-text">
                    ⚠️ 경고 ⚠️<br>
                    절대로 마우스를 조작하지 마세요!<br>
                    작업이 완료될 때까지 기다려주세요.
                </div>
            """, unsafe_allow_html=True)
            
            # 구분선
            st.markdown("<hr>", unsafe_allow_html=True)
            
            # 작업 실행 (placeholder들을 인자로 전달)
            with st.spinner(""):
                run_macro(selected_monitor, working_text, warning_text, status_container)

if __name__ == "__main__":
    main()
