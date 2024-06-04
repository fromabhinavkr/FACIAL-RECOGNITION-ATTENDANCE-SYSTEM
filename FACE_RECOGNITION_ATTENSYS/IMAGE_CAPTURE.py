import cv2

cam_port = 0
cam = cv2.VideoCapture(cam_port)

inp = input('Enter person name: ')

while True:
    ret, image = cam.read()
    cv2.imshow(inp, image)

    key = cv2.waitKey(1)
    if key == ord('q'):
        cv2.imwrite(f"{inp}.png", image)
        print("Image taken and saved.")
        break
    elif key == ord('r'):
        print("Resetting...")
        continue
    elif key == ord('x'):
        print("Exiting...")
        break

cam.release()
cv2.destroyAllWindows()