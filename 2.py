import cv2

# Function to get mouse click coordinates
def get_coordinates(event, x, y, flags, param):
    if event == cv2.EVENT_LBUTTONDOWN:
        print(f"Clicked at pixel coordinates (x, y): ({x}, {y})")

# Load the image
image_path = 'face_recognition/Resources/background.png'
image = cv2.imread(image_path)

# Check if the image is loaded successfully
if image is not None:
    # Create a window and set the callback function
    cv2.namedWindow('Image')
    cv2.setMouseCallback('Image', get_coordinates)

    # Display the image
    cv2.imshow('Image', image)

    # Wait for the user to click and retrieve coordinates
    cv2.waitKey(0)

    # Close the window
    cv2.destroyAllWindows()
else:
    print("Error loading the image.")
