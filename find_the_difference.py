# モジュール読み込み
import cv2


def find_the_diff(before_image_path, after_image_path):
    img1 = cv2.imread('image_file/before_01.png')
    img2 = cv2.imread('image_file/after_01.png')
    img_base = cv2.imread('image_file/before_01.png')
    img_base_increment = cv2.imread('image_file/before_01.png')

    # 変更前-変更後
    img_diff = cv2.subtract(img2, img1)
    for x in range(img_diff.shape[0]):
        for y in range(img_diff.shape[1]):
            b, g, r = img_diff[x, y]
            # 対象ピクセルに差がある場合
            if (b, g, r) != (0, 0, 0):
                img_base[x, y] = b, 0, 0
    cv2.imwrite('diff_result_decrement.png', img_base)

    # 変更後-変更前
    img_diff2 = cv2.subtract(img1, img2)
    for x in range(img_diff2.shape[0]):
        for y in range(img_diff2.shape[1]):
            b, g, r = img_diff2[x, y]
            # 対象ピクセルに差がある場合
            if (b, g, r) != (0, 0, 0):
                img_base[x, y] = 0, 0, r
                img_base_increment[x, y] = 0, 0, r
    cv2.imwrite('diff_result_increment.png', img_base_increment)
    cv2.imwrite('diff_result_marge.png', img_base)
