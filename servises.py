import os
import pyautogui
import time
import cv2
import pytesseract
import openpyxl
from PIL import Image
from settings import *

pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"


def find_png(picture, button):
    retry_counter = 0
    while retry_counter < 8:
        try:
            x, y = pyautogui.locateCenterOnScreen(picture)
            pyautogui.moveTo(x, y, 0.2)
            pyautogui.click(button=button)
            print(f'нашел с {retry_counter + 1} попытки изображение {picture}')
            retry_counter = 10  # to break the loop
        except:
            time.sleep(1)  # retry after some time, i.e. 1 sec
            retry_counter += 1


def find_png_confidence(picture, button):
    retry_counter = 0
    while retry_counter < 8:
        try:
            x, y = pyautogui.locateCenterOnScreen(picture, confidence=0.9)
            if x > 0:
                pyautogui.moveTo(x, y, 0.2)
                pyautogui.click(button=button)
                print(f'нашел с {retry_counter + 1} попытки изображение {picture}')
                retry_counter = 10  # to break the loop
        except:
            time.sleep(1)  # retry after some time, i.e. 1 sec
            retry_counter += 1


def get_file(folder):
    cur_dir = Path.cwd()
    dir = Path(cur_dir, folder)
    files = Path(dir).glob('*.png')
    return files


def tile_screenshot(filename, dir_in, dir_out, count_row):
    name, ext = os.path.splitext(filename)
    img = Image.open(os.path.join(dir_in, filename))

    # параметры для базы
    rows = count_row
    row_start_height = 162
    row_height = 21
    row_start_width = 1023
    row_width = 350
    shift = 1

    for row in range(1, rows + 1):
        out = os.path.join(dir_out, f'{name}_row_{row}{ext}')
        box = (row_start_width, row_start_height, row_start_width + row_width, row_start_height + row_height)
        img.crop(box).save(out)
        row_start_height += row_height - shift


def ocr_png(file):
    img = cv2.imread(file)
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    img = cv2.resize(img, (0, 0), fx=5, fy=5)

    # # Пороговое значение изображения по методу бинаризации Оцу
    # img = cv2.GaussianBlur(img, (5, 5), 0)
    # retval, img = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)

    # Адаптивное пороговое значение для изображения
    img = cv2.medianBlur(img, 5)
    img = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 17, 2)
    img = cv2.medianBlur(img, 5)
    img = cv2.GaussianBlur(img, (3, 3), 0)

    # cv2.imshow('asd', img)
    # cv2.waitKey(0)
    # cv2.destroyAllWindows()
    custom_config = r'-l rus+eng --psm 7 --oem 3'
    text = pytesseract.image_to_string(img, config=custom_config)
    return text


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
    find_png_confidence("steps/oper_na_osnove.png", "left")
    find_png_confidence("steps/comment.png", "left")
    pyautogui.write('123')
    find_png_confidence("steps/proverka.png", "left")
    find_png_confidence("steps/ring.png", "left")
    find_png_confidence("steps/save.png", "left")
    find_png_confidence("steps/without_reserve.png", "left")
    find_png_confidence("steps/save_in_bufer.png", "left")
    pyautogui.move(0, 50, 0.3)
    pyautogui.click()
    # find_png_confidence("steps/register_complete.png", "left", 0.2) # не хочет искать эту картинку
    pyautogui.move(195, 5, 0.5)
    pyautogui.click()
    time.sleep(2.5)
    print(adr_file, '----', 'ОК')


def run_move_to_row(file):
    path_file = str(Path(dir_order, file))
    find_png(path_file, "left")


def loging_xlsx(file):
    file_name = Path(file).stem
    wb = openpyxl.load_workbook("Книга1.xlsx")
    sh = wb["Лист1"]
    for row in sh:
        for cell in row:
            if str(file_name.split("_")[0]) == str(cell.value):
                sh.cell(row=cell.row, column=18).value = "ОК"
                wb.save("Книга1.xlsx")
                return True


def run_convert_to_operation_save_with_reserve(file):
    path_file = str(Path(dir_out, file))
    x, y = pyautogui.locateCenterOnScreen(path_file)
    pyautogui.moveTo(x, y, 0.3)
    pyautogui.click(button='right')
    time.sleep(1)
    find_png_confidence("steps/preobr_tov_oper.png", "left")
    # find_png_confidence("steps/comment.png", "left")
    # pyautogui.write('123')
    find_png_confidence("steps/proverka.png", "left")
    find_png_confidence("steps/ring.png", "left")
    find_png_confidence("steps/save.png", "left")
    find_png_confidence("steps/with_reserve.png", "left")
    find_png_confidence("steps/save_in_bufer_old_oper.png", "left")
    pyautogui.move(0, 35, 0.4)
    pyautogui.click()
    # find_png_confidence("steps/register_complete.png", "left") # не хочет искать эту картинку
    pyautogui.move(195, 5, 0.5)
    pyautogui.click()
    time.sleep(2)
    print(path_file, '----', 'ОК')


def run_register_operation(file):
    path_file = str(Path(dir_out, file))
    # x, y = pyautogui.locateCenterOnScreen(path_file)
    # pyautogui.moveTo(x, y, 0.3)
    # pyautogui.click(button='right')
    find_png(path_file, "right")
    # time.sleep(1)
    find_png_confidence("steps/preobr_tov_oper.png", "left")
    # find_png_confidence("steps/comment.png", "left")
    # pyautogui.write('123')
    find_png_confidence("steps/proverka.png", "left")
    find_png_confidence("steps/get_number_for_oper.png", "left")
    find_png_confidence("steps/get_number_auto.png", "left")
    find_png_confidence("steps/get_number_question.png", "left")
    pyautogui.move(0, 35, 0.3)
    pyautogui.click()
    find_png_confidence("steps/ring.png", "left")
    find_png_confidence("steps/register.png", "left")
    # time.sleep(1) # отключил на время теста
    find_png_confidence("steps/register_question.png", "left")
    pyautogui.move(0, 35, 0.4)
    pyautogui.click()
    # time.sleep(0.3) # отключил на время теста
    # find_png_confidence("steps/register_complete.png", "left", 0.3) # не хочет искать эту картинку
    pyautogui.move(195, 5, 0.7)
    pyautogui.click()
    # time.sleep(2) # отключил на время теста
    print(path_file, '----', 'ОК')


def run_edit_order(file):
    path_file = str(Path(dir_out, file))
    find_png(path_file, "right")
    # time.sleep(1)
    find_png_confidence("steps/edit.png", "left")
    find_png_confidence("steps/number_doc.png", "left")
    pyautogui.move(190, 0, 0.2)
    pyautogui.click()
    # time.sleep(0.2)
    pyautogui.press('backspace', presses=15)
    pyautogui.write(str(file.name.split('.')[0].split('_')[0]))
    find_png_confidence("steps/number_doc.png", "left")
    pyautogui.move(670, 445, 0.2)
    pyautogui.click()
    find_png_confidence("steps/save_edit.png", "left")
    pyautogui.move(-10, 50, 0.2)
    pyautogui.click()
    find_png_confidence("steps/change_complete.png", "left")
    # time.sleep(5)
    pyautogui.move(55, 50, 0.2)
    pyautogui.click()
    # time.sleep(1)
