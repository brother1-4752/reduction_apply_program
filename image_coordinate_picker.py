import cv2

# =========================
# PDF A4 크기 (point 단위)
# =========================
PDF_WIDTH = 595
PDF_HEIGHT = 842

IMAGE_PATH = "assets/form_bg.png"

img = cv2.imread(IMAGE_PATH)

if img is None:
    raise FileNotFoundError("이미지를 찾을 수 없습니다.")

# A4 비율로 리사이즈 (왜곡 없음, 동일 비율)
img_resized = cv2.resize(img, (PDF_WIDTH, PDF_HEIGHT))

print("PDF 좌표 추출 모드")
print("왼쪽 아래가 (0,0)")
print("ESC 누르면 종료")


def mouse_callback(event, x, y, flags, param):
    if event == cv2.EVENT_LBUTTONDOWN:

        # OpenCV는 왼쪽 상단이 (0,0)
        # PDF는 왼쪽 하단이 (0,0)
        pdf_x = x
        pdf_y = PDF_HEIGHT - y

        print(f"PDF 좌표: ({pdf_x}, {pdf_y})")

        cv2.circle(display, (x, y), 4, (0, 0, 255), -1)
        cv2.imshow("PDF Coordinate Picker", display)


display = img_resized.copy()

cv2.namedWindow("PDF Coordinate Picker")
cv2.setMouseCallback("PDF Coordinate Picker", mouse_callback)
cv2.imshow("PDF Coordinate Picker", display)

while True:
    key = cv2.waitKey(1)
    if key == 27:  # ESC
        break

cv2.destroyAllWindows()
