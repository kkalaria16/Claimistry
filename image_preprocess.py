import cv2
import numpy as np
from pathlib import Path

def order_points(pts):
    """Order points in top-left, top-right, bottom-right, bottom-left order."""
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]
    return rect

def four_point_transform(image, pts):
    """Apply perspective transform to straighten the image."""
    rect = order_points(pts)
    (tl, tr, br, bl) = rect
    widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
    widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
    maxWidth = max(int(widthA), int(widthB))
    heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
    heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
    maxHeight = max(int(heightA), int(heightB))
    dst = np.array([[0, 0], [maxWidth - 1, 0], [maxWidth - 1, maxHeight - 1], [0, maxHeight - 1]], dtype="float32")
    M = cv2.getPerspectiveTransform(rect, dst)
    warped = cv2.warpPerspective(image, M, (maxWidth, maxHeight))
    return warped

def clean_document_effect(image: np.ndarray, block_size: int = 35, c: int = 10) -> np.ndarray:
    """
    1. Convert to gray and blur a little.
    2. Adaptive‐mean threshold to separate paper vs. ink.
    3. Whiten all 'paper' pixels in the original color image.
    """
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    thresh = cv2.adaptiveThreshold(
        blur, 255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY,
        block_size,
        c
    )
    out = image.copy()
    paper_mask = (thresh == 255)
    out[paper_mask] = (255, 255, 255)
    return out

def autocrop_image(image_path, output_dir):
    """Detect and crop both header and expenses tables from the image, then combine them."""
    image = cv2.imread(image_path)
    if image is None:
        raise ValueError(f"Could not load image at {image_path}")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(blurred, 75, 200)
    contours, _ = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        raise ValueError("No contours found in the image")
    contours = sorted(contours, key=cv2.contourArea, reverse=True)[:2]
    table_images = []
    table_positions = []
    for contour in contours:
        epsilon = 0.02 * cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, epsilon, True)
        if len(approx) == 4:
            warped = four_point_transform(image, approx.reshape(4, 2))
            x, y, w, h = cv2.boundingRect(contour)
            table_images.append(warped)
            table_positions.append(y)
        else:
            raise ValueError("Could not detect a quadrilateral (4 corners) for a table")
    combined = [img for _, img in sorted(zip(table_positions, table_images), key=lambda x: x[0])]
    widths = [img.shape[1] for img in combined]
    max_width = max(widths)
    resized_tables = [cv2.resize(img, (max_width, int(img.shape[0] * max_width / img.shape[1]))) for img in combined]
    enhanced_tables = []
    border_color = (0, 128, 255)
    for img in resized_tables:
        enhanced = clean_document_effect(img)
        bordered = cv2.copyMakeBorder(
            enhanced,
            top=8, bottom=8, left=8, right=8,
            borderType=cv2.BORDER_CONSTANT,
            value=border_color
        )
        enhanced_tables.append(bordered)
    gap_height = 30
    white_gap = np.full((gap_height, enhanced_tables[0].shape[1], 3), (255, 255, 255), dtype=np.uint8)
    final_image = cv2.vconcat([enhanced_tables[0], white_gap, enhanced_tables[1]])
    output_path = output_dir / f"combined_{Path(image_path).name}"
    cv2.imwrite(str(output_path), final_image)
    return output_path
