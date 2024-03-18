import time
import pyautogui

time.sleep(3)
while True:
    try:
        x, y = pyautogui.locateCenterOnScreen('autogui/button.PNG')
        pyautogui.click(x, y)
        break
    except:
        pass
print("foi")