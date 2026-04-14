import glob
import logging
import os
import time
from dataclasses import dataclass
from typing import List, Optional, Tuple

import cv2
import keyboard
import mss
import numpy as np
import ctypes
try:
    import win32gui
except ImportError:
    win32gui = None

import win32api
import win32con

# DPI Awareness
try:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass

from config import CONFIG


@dataclass
class Template:
    name: str
    image: np.ndarray


def setup_logging() -> None:
    os.makedirs("logs", exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler("logs/runtime.log", encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def normalize_poll_interval(interval: float) -> float:
    if interval <= 0:
        logging.warning("poll_interval_sec <= 0, fallback to 5.0")
        return 5.0
    if interval > 5.0:
        logging.warning("poll_interval_sec > 5.0, clamped to 5.0")
        return 5.0
    return interval


def find_window_by_keyword(keyword: str) -> Optional[int]:
    if win32gui is None:
        # Mock for non-Windows testing (e.g. testing logic on Linux)
        return 1
    keyword_lc = keyword.lower()
    result_hwnd: Optional[int] = None

    def _enum_handler(hwnd: int, _ctx: object) -> None:
        nonlocal result_hwnd
        if result_hwnd is not None:
            return
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return
        if keyword_lc in title.lower():
            result_hwnd = hwnd

    win32gui.EnumWindows(_enum_handler, None)
    return result_hwnd


def get_client_rect_on_screen(hwnd: int) -> Tuple[int, int, int, int]:
    if win32gui is None:
        # Mock for non-Windows testing
        return 0, 0, CONFIG.expected_window_width, CONFIG.expected_window_height
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    client_w = right - left
    client_h = bottom - top
    screen_left, screen_top = win32gui.ClientToScreen(hwnd, (0, 0))
    return screen_left, screen_top, client_w, client_h


def build_roi(left: int, top: int, width: int, height: int) -> Tuple[int, int, int, int]:
    roi_left = left + int(width * CONFIG.roi_left_ratio)
    roi_top = top + int(height * CONFIG.roi_top_ratio)
    roi_width = int(width * CONFIG.roi_width_ratio)
    roi_height = int(height * CONFIG.roi_height_ratio)
    return roi_left, roi_top, max(roi_width, 1), max(roi_height, 1)


def preprocess(image_bgr: np.ndarray) -> np.ndarray:
    # Always convert to gray
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    
    # If using edge match, perform Canny, otherwise simple blur is enough
    if CONFIG.use_edge_match:
        # User observation: edges might be too aggressive for simple buttons
        # Try a more standard Canny for consistency
        return cv2.Canny(gray, 100, 200)
    
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    return gray


def load_templates() -> List[Template]:
    pattern = os.path.join(CONFIG.template_dir, CONFIG.template_pattern)
    paths = sorted(glob.glob(pattern))
    templates: List[Template] = []

    for path in paths:
        raw = cv2.imread(path)
        if raw is None:
            logging.warning("skip unreadable template: %s", path)
            continue
        # Use simple gray for yes.png if it's a simple button
        if "yes" in path.lower():
             processed = cv2.cvtColor(raw, cv2.COLOR_BGR2GRAY)
        else:
             processed = preprocess(raw)
        templates.append(Template(name=os.path.basename(path), image=processed))

    if not templates:
        raise FileNotFoundError(
            "No template images found. Put PNG files into templates/ first."
        )

    logging.info("Loaded %d templates", len(templates))
    return templates


def capture_bgr(sct: mss.mss, roi: Tuple[int, int, int, int]) -> np.ndarray:
    left, top, width, height = roi
    # Note: MSS grab uses SCREEN coordinates.
    monitor = {"left": left, "top": top, "width": width, "height": height}
    try:
        shot = np.array(sct.grab(monitor), dtype=np.uint8)
        # Ensure we drop alpha channel or handle BGRA correctly
        return shot[:, :, :3]
    except Exception as e:
        if win32gui is None:
            # On Linux/No-display environment, return dummy if file exists
            if os.path.exists("template.png"):
                dummy = cv2.imread("template.png")
                # Crop same ROI logic to simulate
                r_left = int(dummy.shape[1] * CONFIG.roi_left_ratio)
                r_top = int(dummy.shape[0] * CONFIG.roi_top_ratio)
                r_w = int(dummy.shape[1] * CONFIG.roi_width_ratio)
                r_h = int(dummy.shape[0] * CONFIG.roi_height_ratio)
                return dummy[r_top:r_top+r_h, r_left:r_left+r_w]
        raise e


def best_match_score(frame_processed: np.ndarray, templates: List[Template], scale: float = 1.0) -> Tuple[float, str, Tuple[int, int]]:
    best_score = -1.0
    best_name = ""
    best_loc = (0, 0)
    fh, fw = frame_processed.shape[:2]

    for tpl in templates:
        tpl_img = tpl.image
        # If running on non-reference resolution, resize template dynamically
        if abs(scale - 1.0) > 0.01:
            new_w = max(1, int(tpl_img.shape[1] * scale))
            new_h = max(1, int(tpl_img.shape[0] * scale))
            tpl_img = cv2.resize(tpl_img, (new_w, new_h), interpolation=cv2.INTER_AREA)

        th, tw = tpl_img.shape[:2]
        if th > fh or tw > fw:
            continue
        result = cv2.matchTemplate(frame_processed, tpl_img, cv2.TM_CCOEFF_NORMED)
        _min_val, max_val, _min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val > best_score:
            best_score = float(max_val)
            best_name = tpl.name
            # Center of the match
            best_loc = (max_loc[0] + tw // 2, max_loc[1] + th // 2)

    return best_score, best_name, best_loc


def press_once(hwnd: int, key: str) -> None:
    if win32gui is None:
        logging.info("Mocking key press for non-Windows environment: %s", key)
        return
    
    # Handle special keys or length > 1
    if key.lower() == "esc":
        vk_code = win32con.VK_ESCAPE
    elif len(key) == 1:
        vk_code = win32api.VkKeyScan(key) & 0xFF
    else:
        logging.warning("Unsupported key string: %s", key)
        return

    # Map virtual key to scan code
    scan_code = win32api.MapVirtualKey(vk_code, 0)
    
    # Use PostMessage for more reliable background input
    lparam_down = 1 | (scan_code << 16)
    lparam_up = 1 | (scan_code << 16) | (1 << 30) | (1 << 31)
    
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lparam_down)
    time.sleep(0.05)  # Brief delay to simulate human press duration
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lparam_up)


def click_at(hwnd: int, x: int, y: int) -> None:
    if win32gui is None:
        logging.info("Mocking click at (%d, %d)", x, y)
        return
    
    # Convert client (x, y) to screen coordinates
    try:
        screen_pos = win32gui.ClientToScreen(hwnd, (x, y))
        # Use win32api to perform a physical mouse click
        win32api.SetCursorPos(screen_pos)
        time.sleep(0.1)
        # mouse_event uses specific flags for down and up
        # LEFTDOWN = 0x0002, LEFTUP = 0x0004
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        logging.info("Performed physical click at screen pos %s", screen_pos)
    except Exception as e:
        logging.warning("Failed to perform physical click: %s", e)


def run() -> None:
    setup_logging()

    logging.info("Starting detector. Stop hotkey: %s", CONFIG.stop_hotkey)
    logging.info("This script is for authorized testing only.")

    print("\n请选择运行模式:")
    print("1: 聚能模式 (自动键入 X)")
    print("2: 逃跑模式 (自动键入 ESC 并点击确认)")
    print("\n[提示] 建议游戏分辨率设为 2560x1600 (16:10) 以获得最佳识别效果。")
    print("如果使用其他比例(如4:3)，图像可能会形变导致识别失败。")
    choice = input("请输入选项 (1 或 2): ").strip()
    mode = "battle" if choice != "2" else "escape"
    logging.info("已选择模式: %s", "聚能模式" if mode == "battle" else "逃跑模式")

    templates = load_templates()
    interval = normalize_poll_interval(CONFIG.poll_interval_sec)

    hit_streak = 0
    miss_streak = 0
    in_battle_state = False
    last_trigger_time = 0.0

    with mss.mss() as sct:
        while True:
            # On Windows, monitor hotkey to stop
            if win32gui is not None:
                if keyboard.is_pressed(CONFIG.stop_hotkey):
                    logging.info("Stop hotkey pressed. Exiting.")
                    break
            
            hwnd = find_window_by_keyword(CONFIG.window_title_keyword)
            if hwnd is None:
                logging.warning("Game window not found: %s", CONFIG.window_title_keyword)
                time.sleep(interval)
                continue

            left, top, width, height = get_client_rect_on_screen(hwnd)
            if width <= 0 or height <= 0:
                logging.warning("Invalid window size: %sx%s", width, height)
                time.sleep(interval)
                continue

            # Calculate dynamic scale based on reference resolution
            scale = width / CONFIG.ref_width
            if abs(scale - 1.0) > 0.05:
                logging.debug("Scaling templates by factor: %.2f (width=%d)", scale, width)

            roi = build_roi(left, top, width, height)
            frame_bgr = capture_bgr(sct, roi)
            frame_processed = preprocess(frame_bgr)

            # Pass the scale factor for resolution independence
            score, name, center_loc = best_match_score(frame_processed, templates, scale=scale)
            is_hit = score >= CONFIG.match_threshold

            if is_hit:
                hit_streak += 1
                miss_streak = 0
            else:
                hit_streak = 0
                miss_streak += 1

            if not in_battle_state:
                detected = hit_streak >= CONFIG.required_hits
            else:
                detected = miss_streak < CONFIG.release_misses

            logging.info(
                "score=%.3f hit=%s hit_streak=%d miss_streak=%d tpl=%s",
                score,
                is_hit,
                hit_streak,
                miss_streak,
                name,
            )

            now = time.time()
            cooldown_ready = (now - last_trigger_time) >= CONFIG.trigger_cooldown_sec
            
            # Action logic based on mode
            if detected:
                # IMPORTANT: Only trigger if we actually matched a template (is_hit)
                # This prevents triggering on "miss_streak" logic when we haven't seen the target
                if is_hit and (now - last_trigger_time >= CONFIG.trigger_cooldown_sec):
                    if mode == "battle":
                        press_once(hwnd, CONFIG.press_key)
                        last_trigger_time = now
                        logging.info("Triggered key: %s (Continuous)", CONFIG.press_key)
                    elif mode == "escape":
                        press_once(hwnd, "esc")
                        logging.info("Triggered Escape")
                        
                        # Capture again to find "Yes" button in popup
                        button_clicked = False
                        # Increase attempts for popup animation
                        for i in range(10):
                            time.sleep(0.3)
                            full_shot = capture_bgr(sct, (left, top, width, height))
                            
                            # Using the dual matching strategy with scaling
                            full_processed_edge = preprocess(full_shot)
                            full_processed_gray = cv2.cvtColor(full_shot, cv2.COLOR_BGR2GRAY)
                            
                            best_score_this_round = -1.0
                            best_loc_this_round = (0, 0)
                            
                            for tpl in templates:
                                if "yes" not in tpl.name.lower(): continue
                                t_img = tpl.image
                                if abs(scale - 1.0) > 0.01:
                                    t_img = cv2.resize(t_img, (max(1, int(t_img.shape[1] * scale)), max(1, int(t_img.shape[0] * scale))), interpolation=cv2.INTER_AREA)

                                res_edge = cv2.matchTemplate(full_processed_edge, t_img, cv2.TM_CCOEFF_NORMED)
                                res_gray = cv2.matchTemplate(full_processed_gray, t_img, cv2.TM_CCOEFF_NORMED)
                                _, max_v_edge, _, max_l_edge = cv2.minMaxLoc(res_edge)
                                _, max_v_gray, _, max_l_gray = cv2.minMaxLoc(res_gray)
                                
                                cur_v, cur_l = (max_v_edge, max_l_edge) if max_v_edge > max_v_gray else (max_v_gray, max_l_gray)
                                if cur_v > best_score_this_round:
                                    best_score_this_round = cur_v
                                    best_loc_this_round = (cur_l[0] + t_img.shape[1]//2, cur_l[1] + t_img.shape[0]//2)

                            if best_score_this_round >= (CONFIG.match_threshold * 0.8):
                                # Perform physical click (which includes logging the screen pos)
                                click_at(hwnd, best_loc_this_round[0], best_loc_this_round[1])
                                button_clicked = True
                                break
                        
                        if not button_clicked:
                            logging.warning("Could not find confirmation button 'yes.png' after ESC")
                        
                        # Use a longer cooldown for escape to prevent ESC spamming while dialog is closing
                        last_trigger_time = now + 2.0 

            in_battle_state = detected
            time.sleep(interval)


if __name__ == "__main__":
    run()
