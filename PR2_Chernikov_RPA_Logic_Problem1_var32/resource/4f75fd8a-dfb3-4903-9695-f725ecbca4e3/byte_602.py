# !/usr/bin/python 
# -*- coding: utf8 -*- 
# Puzzle RPA version: 2.3.0 
# remote
def downRange(start, stop, step):

    while start >= stop:
        yield start
        start -= abs(step)
def upRange(start, stop, step):

    while start <= stop:
        yield start
        start += abs(step)
import sys

sys.dont_write_bytecode = True

# storage:
from puzzle_logger import configure_logger, log_process, send_message_websocket
from trace_utils import format_traceback
from pathlib import Path

import convert_date_mod
import excel_update_row
import files_and_folders
import interaction_with_excel_data
import read_from_excel
import str_to_int
import user_notice_2

# generated

logger = configure_logger()
puzzle_logger_path = Path(__file__).absolute()
logger.info(f'Старт робота: {puzzle_logger_path}')

if __name__ == "__main__":
    def main_byte_602_proc():
        try:
            log_process(window_log=True,block_text='Группировка блоков')
            # Инициализация переменных
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #q5tx|+$BkH$]q]|ST$L!
            total_bookings = 0
            log_process(block_text='Присвоить значение переменной')
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #br@]/f2Iz=2!(*y5.)b2
            confirmed_bookings = 0
            log_process(block_text='Присвоить значение переменной')
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #o(8WUA:7Lh,_Uy5bXylS
            cancelled_bookings = 0
            log_process(block_text='Присвоить значение переменной')
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #~nX[NW0Ne}K%2kM!0+iR
            completed_sessions = 0
            log_process(block_text='Присвоить значение переменной')
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #u!Wju-6RzxK?vk?^[gC*
            no_show_cases = 0
            log_process(block_text='Присвоить значение переменной')
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #YE!YRGof1g6EMG_K{W(g
            short_notice_cancel = 0
            log_process(block_text='Присвоить значение переменной')
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #}6S#J{n_2xqtenjh`GyB
            total_revenue_loss = 0
            log_process(block_text='Присвоить значение переменной')
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #Ir/e?Dr~i^{IL8Ts$.#p
            one_day = (convert_date_mod.convert_date_mod('02.02.2002','datetime_object',block_text="Преобразовать дату",window_log=True, current_language="ru")) - (convert_date_mod.convert_date_mod('01.02.2002','datetime_object',block_text="Преобразовать дату",window_log=True, current_language="ru"))
            log_process(block_text='Присвоить значение переменной')
            

            log_process(block_text='Группировка блоков')

            # Присваивает переменной значение вставки
            log_process(window_log=True,block_text='Присвоить значение переменной')
            #lYU1)-LCL|mxv7YJlpFP
            table = read_from_excel.read_from_excel_2((files_and_folders.get_executable_path('GymBookings_Q1.xlsx',binary_path='TRUE',block_text="Относительный путь",window_log=True, current_language="ru")),'Bookings_Q1',True,dtype=None,block_text="Прочитать из Excel",window_log=True, current_language="ru")

            log_process(block_text='Присвоить значение переменной')

            log_process(window_log=True,block_text='Цикл по (от до с шагом)')
            #[)yJHq-%M@]QF.3=XdsM
            i_end = len(table) - 1
            for i in (1 <= i_end) and upRange(1, i_end, 1) or downRange(1, i_end, 1):
                log_process(window_log=True,block_text='Группировка блоков')
                # Получение столбцов строки
                log_process(window_log=True,block_text='Присвоить значение переменной')
                #Iie5tsiGPD}5x:p}t]a.
                BookingID = (interaction_with_excel_data.get_excel_cell_data(data=table,column='A',index=i,block_text="Получить значение ячейки табличных данных",window_log=True, current_language="ru")
                ).replace(' ', '').upper()
                log_process(block_text='Присвоить значение переменной')
                log_process(window_log=True,block_text='Присвоить значение переменной')
                #sw;jB4V]$G+z%QfXB.SD
                GymCode = (interaction_with_excel_data.get_excel_cell_data(data=table,column='B',index=i,block_text="Получить значение ячейки табличных данных",window_log=True, current_language="ru")
                ).replace(' ', '').upper()
                log_process(block_text='Присвоить значение переменной')
                log_process(window_log=True,block_text='Присвоить значение переменной')
                #51#J@Y2WON0%^zwdguvE
                BookingDate = interaction_with_excel_data.get_excel_cell_data(data=table,column='E',index=i,block_text="Получить значение ячейки табличных данных",window_log=True, current_language="ru")
                log_process(block_text='Присвоить значение переменной')
                log_process(window_log=True,block_text='Присвоить значение переменной')
                #nB=wf}$@MZ/=wJcwR.Eu
                SessionDate = interaction_with_excel_data.get_excel_cell_data(data=table,column='F',index=i,block_text="Получить значение ячейки табличных данных",window_log=True, current_language="ru")
                log_process(block_text='Присвоить значение переменной')
                log_process(window_log=True,block_text='Присвоить значение переменной')
                #B%n5zw_/Z#ll-MPeb|mI
                BookingStatus = interaction_with_excel_data.get_excel_cell_data(data=table,column='G',index=i,block_text="Получить значение ячейки табличных данных",window_log=True, current_language="ru")
                log_process(block_text='Присвоить значение переменной')
                log_process(window_log=True,block_text='Присвоить значение переменной')
                #em%b6OQhxmNcAOftNcA~
                TotalPrice = interaction_with_excel_data.get_excel_cell_data(data=table,column='H',index=i,block_text="Получить значение ячейки табличных данных",window_log=True, current_language="ru")
                log_process(block_text='Присвоить значение переменной')
                log_process(window_log=True,block_text='Присвоить значение переменной')
                #*ryVtUmd71J[f%w(z`@=
                DurationHours = interaction_with_excel_data.get_excel_cell_data(data=table,column='I',index=i,block_text="Получить значение ячейки табличных данных",window_log=True, current_language="ru")
                log_process(block_text='Присвоить значение переменной')
                

                log_process(block_text='Группировка блоков')

                log_process(window_log=True,block_text='Группировка блоков')
                # Расчеты
                log_process(window_log=True,block_text='Присвоить значение переменной')
                #{~_fe4]oZJa;e,c^`]Ee
                days_before_session = SessionDate - BookingDate
                log_process(block_text='Присвоить значение переменной')
                log_process(window_log=True,block_text='Если – выполнить')
                #)B6xW/k=,AqEZ8WD`w8%
                if BookingStatus == 'Cancelled':
                    log_process(window_log=True,block_text='Присвоить значение переменной')
                    #Ni,q8GoOvW2E=.ThSA2r
                    potential_loss = TotalPrice

                    log_process(block_text='Присвоить значение переменной')

                else:
                    log_process(window_log=True,block_text='Присвоить значение переменной')
                    #lt^fis~Htu+hU=L21.h%
                    potential_loss = 0

                    log_process(block_text='Присвоить значение переменной')

                log_process(block_text='Если – выполнить')
                

                log_process(block_text='Группировка блоков')

                log_process(window_log=True,block_text='Группировка блоков')
                # Логика классификации
                log_process(window_log=True,block_text='Если – выполнить')
                #nx%HLLhRXq;O_jzHWSKV
                if BookingStatus == 'Confirmed':
                    log_process(window_log=True,block_text='Присвоить значение переменной')
                    #-bcH`/;Uk1oBF#3h5.)o
                    confirmed_bookings = confirmed_bookings + 1

                    log_process(block_text='Присвоить значение переменной')

                log_process(block_text='Если – выполнить')
                log_process(window_log=True,block_text='Если – выполнить')
                #90K^`(9^4R53V21tkT[k
                if BookingStatus == 'Cancelled':
                    log_process(window_log=True,block_text='Присвоить значение переменной')
                    #Fa=UePlz6#~2RD/;D^[+
                    cancelled_bookings = cancelled_bookings + 1

                    log_process(block_text='Присвоить значение переменной')

                log_process(block_text='Если – выполнить')
                log_process(window_log=True,block_text='Если – выполнить')
                #PDl8Cl_E]1m`Mw`JZ;+?
                if BookingStatus == 'Completed':
                    log_process(window_log=True,block_text='Присвоить значение переменной')
                    #xEh5|]myDs2EMaBi+H-$
                    completed_sessions = completed_sessions + 1

                    log_process(block_text='Присвоить значение переменной')

                log_process(block_text='Если – выполнить')
                log_process(window_log=True,block_text='Если – выполнить')
                #leN,k`e*9/h|KZR[jL7K
                if BookingStatus == 'NoShow':
                    log_process(window_log=True,block_text='Присвоить значение переменной')
                    #X,pGAteld|mSZ_g1K9/i
                    no_show_cases = no_show_cases + 1

                    log_process(block_text='Присвоить значение переменной')

                log_process(block_text='Если – выполнить')
                log_process(window_log=True,block_text='Если – выполнить')
                #uZ~Lk`f#z`DC~;uYdVDh
                if BookingStatus == 'Cancelled' and days_before_session <= one_day:
                    log_process(window_log=True,block_text='Присвоить значение переменной')
                    #qr;/6c-q;7qvPpjL!G{(
                    short_notice_cancel = short_notice_cancel + 1

                    log_process(block_text='Присвоить значение переменной')

                log_process(block_text='Если – выполнить')
                log_process(window_log=True,block_text='Если – выполнить')
                #-@-GQhk=ex}`UMadcc1`
                if BookingStatus == 'Cancelled':
                    log_process(window_log=True,block_text='Присвоить значение переменной')
                    #iyLZYrTjQ9mIRTFm#Hej
                    total_revenue_loss = total_revenue_loss + potential_loss

                    log_process(block_text='Присвоить значение переменной')

                log_process(block_text='Если – выполнить')
                

                log_process(block_text='Группировка блоков')

                #Ik9rd_6%,t^WyVbn$c.,
                excel_update_row.add_row(path_file=(files_and_folders.get_executable_path('GymBookings_Result.xlsx',binary_path='TRUE',block_text="Относительный путь",window_log=True, current_language="ru")),sheet_name='Analysis_Q1',row_id=(i - 1),arr=[BookingID, GymCode, (str_to_int.str_to_int(((convert_date_mod.convert_date_mod(SessionDate,'%Y-%m-%d',block_text="Преобразовать дату",window_log=True, current_language="ru"))[8 : 10]),block_text="Преобразовать строку в число",window_log=True, current_language="ru")) - (str_to_int.str_to_int(((convert_date_mod.convert_date_mod(BookingDate,'%Y-%m-%d',block_text="Преобразовать дату",window_log=True, current_language="ru"))[8 : 10]),block_text="Преобразовать строку в число",window_log=True, current_language="ru")), potential_loss, BookingStatus],block_text="Модифицировать строку в Excel",window_log=True, current_language="ru")


            log_process(block_text='Цикл по (от до с шагом)')

            #k|#qZNuN~~@23S/kG6Tr
            user_notice_2.user_notice((f"Анализ бронирований Q1 завершён.\n\nВсего бронирований: {total_bookings}\nПодтверждено: {confirmed_bookings}\nОтменено: {cancelled_bookings}\nЗавершённых занятий: {completed_sessions}\nNo-show случаев: {no_show_cases}\n\nОтмены в последний момент: {short_notice_cancel}\nПотенциальные потери выручки: {total_revenue_loss}\n\nОтчёт сохранён в GymBookings_Result.xlsx (лист Analysis_Q1)."),None,block_text="Уведомление пользователя",window_log=True, current_language="ru")


            logger.info(f'Завершение работы робота: {puzzle_logger_path}')
            send_message_websocket(message_type="python_end")
        except Exception as error_puzzle:
            logger.error(f'{puzzle_logger_path} ' + f'Ошибка: {error_puzzle}')
            error_puzzle_format=format_traceback(error_puzzle)
            send_message_websocket(message_type="python_error", message=error_puzzle_format)
            raise Exception(error_puzzle)
    main_byte_602_proc()