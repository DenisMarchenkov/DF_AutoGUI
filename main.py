from servises import *

if __name__ == '__main__':

    pyautogui.screenshot('in/screenshot.png')
    tile_screenshot('screenshot.png', dir_in, dir_out, count_row=40)

    i = 1
    for file in get_file("out"):
        file_str = str(file)
        name = ocr_png(file_str)
        name = name[:-1].split()[0]
        old_name = os.path.join(dir_out, file)
        new_name = os.path.join(dir_order, name + "_" + str(i) + ".png")
        os.rename(old_name, new_name)
        i += 1

    # создать список из значений первых пяти символов из столбца 'column' книги 'path'
    list_orders_from_excel = [i[:5] for i in get_dataframe(path="Книга1.xlsx", column=8)]

    for file in get_file("orders"):
        order = file.name.split('_')[0]
        if order not in list_orders_from_excel:
            os.remove(file)
        else:
            # run_operation_based_on_row_save_in_buffer_without_reserve(file)
            # run_convert_to_operation_save_with_reserve(file)
            # run_move_to_row(file)
            # run_register_operation(file)
            run_move_to_row(file)
            # run_edit_order(file)
            loging_xlsx(file)

    movement_files(dir_order, dir_order_complete, 'png')
