from time import time

import xlart

if __name__ == '__main__':
    start = time()
    xlart.resize_image('picture.jpg', 'picture_resized.jpg', 200, 200)
    xlart.image_to_xlsx('picture_resized.jpg', 'demo.xlsx')
    end = time()
    print('An xlart file has been generated in ' + str(end - start) + 's.')
