import os
import pyautogui
import time
from pathlib import Path
import cv2
import pytesseract

from PIL import Image
from itertools import product

pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"


def find_png(picture, button, sleep):
    x, y = pyautogui.locateCenterOnScreen(picture)
    pyautogui.moveTo(x, y, 0.2)
    pyautogui.click(button=button)
    time.sleep(sleep)


def find_png_confidence(picture, button, sleep):
    x, y = pyautogui.locateCenterOnScreen(picture, confidence=0.9)
    pyautogui.moveTo(x, y, 0.2)
    pyautogui.click(button=button)
    time.sleep(sleep)


def get_file(folder):
    cur_dir = Path.cwd()
    dir = Path(cur_dir, folder)
    files = Path(dir).glob('*.png')
    return files


def tile(filename, dir_in, dir_out, d):
    im2 = pyautogui.screenshot('in/screenshot.png')
    name, ext = os.path.splitext(filename)
    img = Image.open(os.path.join(dir_in, filename))
    w, h = img.size

    grid = product(range(0, h - h % d, d), range(0, w - w % d, d))
    for i, j in grid:
        box = (j, i, j + d, i + d)
        out = os.path.join(dir_out, f'{name}_{i}_{j}{ext}')
        img.crop(box).save(out)


def tile_screenshot(filename, dir_in, dir_out, count_row):
    name, ext = os.path.splitext(filename)
    img = Image.open(os.path.join(dir_in, filename))
    # w, h = img.size
    # print(w, h)

    # параметры для базы
    rows = count_row
    row_start_height = 162
    row_height = 21
    row_start_width = 1023
    row_width = 200
    shift = 1

    for row in range(1, rows + 1):
        out = os.path.join(dir_out, f'{name}_row_{row}{ext}')
        box = (row_start_width, row_start_height, row_start_width + row_width, row_start_height + row_height)
        img.crop(box).save(out)
        row_start_height += row_height - shift


def ocr_png(file):
    img = cv2.imread(file)
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    # cv2.imshow('Result', img)
    # cv2.waitKey(0)
    config = r'--oem 3 --psm 6'
    name = pytesseract.image_to_string(img, config=config, lang='rus')
    return name


def movement_files(file_source, file_destination, file_extension):
    import os
    get_files = os.listdir(file_source)
    files = filter(lambda x: x.endswith(file_extension), get_files)
    for file in files:
        os.replace(os.path.join(file_source, file), os.path.join(file_destination, file))


def get_dataframe(path, column):
    from pandas import read_excel
    dataframe = read_excel(path, header=None, na_values='не указан')
    dataframe = dataframe[1:]
    dataframe = dataframe.fillna('нет данных')
    dataframe = dataframe[column].tolist()
    return dataframe


def run_operation_based_on_row_save_in_buffer_without_reserve(file):
    adr_file = str(Path(dir_out, file))
    x, y = pyautogui.locateCenterOnScreen(adr_file)
    pyautogui.moveTo(x, y, 0.2)
    pyautogui.click(button='right')
    time.sleep(1)
    find_png_confidence("steps/oper_na_osnove.png", "left", 4)
    find_png_confidence("steps/comment.png", "left", 0.2)
    pyautogui.write('123')
    find_png_confidence("steps/proverka.png", "left", 0.2)
    find_png_confidence("steps/ring.png", "left", 0.2)
    find_png_confidence("steps/save.png", "left", 0.2)
    find_png_confidence("steps/without_reserve.png", "left", 0.2)
    find_png_confidence("steps/save_in_bufer.png", "left", 0.2)
    pyautogui.move(0, 50, 0.3)
    pyautogui.click()
    # find_png_confidence("steps/register_complete.png", "left", 0.2) # не хочет искать эту картинку
    pyautogui.move(195, 5, 0.5)
    pyautogui.click()
    time.sleep(2.5)
    print(adr_file, '----', 'ОК')


def run_move_to_row(file):
    path_file = str(Path(dir_order, file))
    x, y = pyautogui.locateCenterOnScreen(path_file)
    pyautogui.moveTo(x, y, 0.2)
    pyautogui.click()


def run_convert_to_operation_save_with_reserve(file):
    path_file = str(Path(dir_out, file))
    x, y = pyautogui.locateCenterOnScreen(path_file)
    pyautogui.moveTo(x, y, 0.3)
    pyautogui.click(button='right')
    time.sleep(1)
    find_png_confidence("steps/preobr_tov_oper.png", "left", 4)
    # find_png_confidence("steps/comment.png", "left", 0.2)
    # pyautogui.write('123')
    find_png_confidence("steps/proverka.png", "left", 0.4)
    find_png_confidence("steps/ring.png", "left", 0.4)
    find_png_confidence("steps/save.png", "left", 0.4)
    find_png_confidence("steps/with_reserve.png", "left", 0.4)
    find_png_confidence("steps/save_in_bufer_old_oper.png", "left", 0.4)
    pyautogui.move(0, 35, 0.4)
    pyautogui.click()
    # find_png_confidence("steps/register_complete.png", "left", 0.2) # не хочет искать эту картинку
    pyautogui.move(195, 5, 0.5)
    pyautogui.click()
    time.sleep(2)
    print(path_file, '----', 'ОК')


def run_register_operation(file):
    path_file = str(Path(dir_out, file))
    x, y = pyautogui.locateCenterOnScreen(path_file)
    pyautogui.moveTo(x, y, 0.3)
    pyautogui.click(button='right')
    time.sleep(1)
    find_png_confidence("steps/preobr_tov_oper.png", "left", 4)
    # find_png_confidence("steps/comment.png", "left", 0.3)
    # pyautogui.write('123')
    find_png_confidence("steps/proverka.png", "left", 0.3)
    find_png_confidence("steps/get_number_for_oper.png", "left", 0.5)
    find_png_confidence("steps/get_number_auto.png", "left", 0.5)
    find_png_confidence("steps/get_number_question.png", "left", 0.5)
    pyautogui.move(0, 35, 0.3)
    pyautogui.click()
    find_png_confidence("steps/ring.png", "left", 0.3)
    find_png_confidence("steps/register.png", "left", 0.3)
    time.sleep(1)
    find_png_confidence("steps/register_question.png", "left", 0.3)
    pyautogui.move(0, 35, 0.4)
    pyautogui.click()
    time.sleep(0.3)
    #find_png_confidence("steps/register_complete.png", "left", 0.3) # не хочет искать эту картинку
    pyautogui.move(195, 5, 0.5)
    pyautogui.click()
    time.sleep(2)
    print(path_file, '----', 'ОК')


def run_edit_order(file):
    path_file = str(Path(dir_out, file))
    x, y = pyautogui.locateCenterOnScreen(path_file)
    pyautogui.moveTo(x, y, 0.3)
    pyautogui.click(button='right')
    time.sleep(1)
    find_png_confidence("steps/edit.png", "left", 0.3)
    find_png_confidence("steps/number_doc.png", "left", 0.3)
    pyautogui.move(190, 0, 0.4)
    # time.sleep(1)
    pyautogui.click()
    time.sleep(1)
    pyautogui.press('backspace', presses=15)
    #pyautogui.write('1')
    pyautogui.write(str(file.name.split('.')[0]))
    find_png_confidence("steps/number_doc.png", "left", 0.3)
    pyautogui.move(670, 445, 0.4)
    pyautogui.click()
    find_png_confidence("steps/save_edit.png", "left", 0.3)
    pyautogui.move(-10, 50, 0.4)
    pyautogui.click()
    find_png_confidence("steps/change_complete.png", "left", 0.1)
    #time.sleep(5)
    pyautogui.move(55, 50, 0.2)
    pyautogui.click()
    time.sleep(1)

if __name__ == '__main__':
    dir_in = Path(Path.cwd(), 'in')
    dir_out = Path(Path.cwd(), 'out')
    dir_order = Path(Path.cwd(), 'orders')
    dir_order_complete = Path(dir_order, 'complete')

    pyautogui.screenshot('in/screenshot.png')
    tile_screenshot('screenshot.png', dir_in, dir_out, count_row=30)

    for file in get_file("out"):
        file_str = str(file)
        name = ocr_png(file_str)
        name = name[:-1].replace('.', '')
        name = name.replace(' ', '_')[:5]
        os.rename(os.path.join(dir_out, file), os.path.join(dir_order, name + ".png"))

    # создать список из значений первых пяти символов из столбца 'column' книги 'path'
    list_orders_from_excel = [i[:5] for i in get_dataframe(path="Книга1.xlsx", column=8)]

    for file in get_file("orders"):
        order = file.name[:5]
        if order not in list_orders_from_excel:
            os.remove(file)

    files = get_file("orders")
    for file in files:
        # run_operation_based_on_row_save_in_buffer_without_reserve(file)
        # run_move_to_row(file)
        # run_convert_to_operation_save_with_reserve(file)
        # run_register_operation(file)
        run_edit_order(file)

    movement_files(dir_order, dir_order_complete, 'png')
