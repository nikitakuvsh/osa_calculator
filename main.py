from tkinter import Tk, filedialog
from openpyxl import Workbook

from make_table import create_table
from merge_data import merge_data


# ==========================
# Цвета терминала
# ==========================

GREEN = "\033[92m"
RED = "\033[91m"
CYAN = "\033[96m"
PURPLE = "\033[95m"
RESET = "\033[0m"


def banner():

    print(PURPLE + r"""
 ██████╗ ███████╗ █████╗      ██████╗ █████╗ ██╗      ██████╗
██╔═══██╗██╔════╝██╔══██╗    ██╔════╝██╔══██╗██║     ██╔════╝
██║   ██║███████╗███████║    ██║     ███████║██║     ██║     
██║   ██║╚════██║██╔══██║    ██║     ██╔══██║██║     ██║     
╚██████╔╝███████║██║  ██║    ╚██████╗██║  ██║███████╗╚██████╗
 ╚═════╝ ╚══════╝╚═╝  ╚═╝     ╚═════╝╚═╝  ╚═╝╚══════╝ ╚═════╝

                 O S A   C A L C U L A T O R     V 0.0.1
    """ + RESET)


def select_excel():

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    file_path = filedialog.askopenfilename(
        title="Выберите Excel",
        filetypes=[
            ("Excel", "*.xlsx")
        ]
    )

    root.destroy()

    return file_path



def main():

    banner()

    print(CYAN + "[+] Инициализация калькулятора..." + RESET)
    print('    Необходимо выбрать файл с выгрузкой из PowerBI')


    source_file = select_excel()


    if not source_file:

        print(
            RED +
            "[ERROR] Файл не выбран. Операция отменена."
            + RESET
        )

        return



    try:

        print(CYAN + "[+] Загружаю Excel..." + RESET)


        wb = Workbook()
        wb.remove(wb.active)


        print(CYAN + "[+] Объединяю данные..." + RESET)


        export_df, dmr_dict = merge_data(
            workbook=wb,
            source_file=source_file
        )


        print(CYAN + "[+] Строю отчёт..." + RESET)


        create_table(
            wb,
            export_df,
            dmr_dict
        )


        wb._sheets = [
            wb["Report"],
            wb["Export"]
        ]


        wb.save("result.xlsx")


        print(
            GREEN +
            "\n[SUCCESS] OSA Calculator завершил работу!"
            "\n[SUCCESS] Файл result.xlsx создан."
            + RESET
        )


    except PermissionError:

        print(
            RED +
            "\n[ERROR] Не удалось сохранить result.xlsx"
            "\n[INFO] Закройте открытый Excel-файл и запустите расчёт снова."
            + RESET
        )


    except Exception as e:

        print(
            RED +
            "\n[CRITICAL ERROR]"
            f"\n{e}"
            + RESET
        )



if __name__ == "__main__":
    main()