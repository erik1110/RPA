#!/usr/bin/python
# coding:utf-8
'''
TEST
'''
import os
import logging
import datetime
import tkinter as tk
from tkinter import messagebox
import pandas as pd

class MyApp(object):
    """
    define the GUI interface
    """
    def __init__(self):
        '''
        set initial UI of labor insurance
        '''
        self.set_log()
        self.root = tk.Tk()
        self.root.title("TEST")
        self.root.geometry('1000x300')
        self.canvas = tk.Canvas(self.root, height=280, width=500)
        self.canvas.pack(side='top')
        self.setup_ui()
        self.logger.info('HELLO WORLD!')

    def set_log(self):
        '''
        set log
        '''
        if not os.path.exists('./screenshot'):
            os.mkdir('./screenshot')
        if not os.path.exists('./log'):
            os.mkdir('./log')
        log_name = 'log/RPA_%Y%m%d_%H%M%S.log'
        logging.basicConfig(level=logging.INFO,
                            filename=datetime.datetime.now().strftime(log_name),
                            filemode='w',
                            format='%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s')
        self.logger = logging.getLogger(log_name)
        self.logger.handlers.clear()

    def setup_ui(self):
        """
        setup UI interface
        """
        self.label_account = tk.Label(self.root, text='帳號:')
        self.label_pwd = tk.Label(self.root, text='密碼:')
        self.label_file_path = tk.Label(self.root, text='輸入資料夾路徑:')
        self.input_account = tk.Entry(self.root, width=100)
        self.input_pwd = tk.Entry(self.root, show='*', width=100)
        self.input_file_path = tk.Entry(self.root, width=100)
        self.login_button = tk.Button(self.root, command=self.run, text="執行", width=10)
        self.quit_button = tk.Button(self.root, command=self.quit, text="退出", width=10)

    def gui_arrang(self):
        """
        setup position of UI
        """
        self.label_account.place(x=60, y=30)
        self.label_pwd.place(x=60, y=70)
        self.label_file_path.place(x=60, y=110)
        self.input_account.place(x=170, y=30)
        self.input_pwd.place(x=170, y=70)
        self.input_file_path.place(x=170, y=110)
        self.login_button.place(x=130, y=190)
        self.quit_button.place(x=270, y=190)

    def check(self):
        """
        check the input of gui interface
        return:
            True
            False
        """
        # check year and month
        self.account = self.input_account.get()
        self.pwd = self.input_pwd.get()
        if len(self.account) == 0 or len(self.pwd) == 0:
            messagebox.showinfo(title='System Alert', message='帳號密碼不得為空！')
            self.logger.warning('帳號密碼為空值！')
            return False

        # check file_path and file
        self.file_dir = self.input_file_path.get()
        if len(self.file_dir) == 0:
            messagebox.showinfo(title='System Alert', message='檔案路徑不得為空！')
            self.logger.warning('檔案路徑為空值！')
            return False
        
        try:
            self.data = pd.read_excel(self.file)
            if len(self.data) == 0:
                messagebox.showinfo(title='System Alert', message='資料為0筆！')
                self.logger.warning('資料為0筆！')
                return False

        except FileNotFoundError:
            messagebox.showinfo(title='System Alert', message='檔案路徑有問題！')
            self.logger.warning(FileNotFoundError, exc_info=True)
            return False

        except ImportError:
            messagebox.showinfo(title='System Alert', message='無法讀取檔案(pandas ImportError)！')
            self.logger.error(ImportError, exc_info=True)
            return False

        except OSError:
            messagebox.showinfo(title='System Alert', message=f'無法讀取檔案！檔案路徑異常或不存在:{self.file_dir}')
            self.logger.error(OSError, exc_info=True)
            return False

        except Exception as error:
            messagebox.showinfo(title='System Alert', message=error)
            self.logger.error(error, exc_info=True)
            return False

        # check driver
        chromedriver_path = os.path.join(self.file_dir, 'chromedriver.exe')
        self.logger.info('Chromedriver檔案路徑:{}'.format(chromedriver_path))
        if os.path.exists(r'{chromedriver_path}'):
            self.logger.warning('當下路徑沒有chromedriver!')
            messagebox.showinfo(title='System Alert', message='當下路徑沒有chromedriver')
            return False

        return True

    def run(self):
        """
        when you click the button of run, it'll execute
        """
        start_time = datetime.datetime.now()
        try:
            if self.check():
                self.logger.info('EXCELLENT!')
            else:
                self.logger.warning('檢查不通過！')
        except Exception as error:
            self.logger.error(error, exc_info=True)
            messagebox.showinfo(title='System Alert', message='執行異常!')
            raise error
        finally:
            end_time = datetime.datetime.now()
            execution_time = (end_time-start_time).seconds
            execution_time_format = str(datetime.timedelta(seconds=execution_time))
            self.logger.info('Total Execution time:{}'.format(execution_time_format))
            messagebox.showinfo(title='System Alert', message=f'執行時間:{execution_time_format}')

    def quit(self):
        """
        when you click the button of quit, it'll execute
        """
        self.root.destroy()

def main():
    """
    main function for MyApp
    """
    # initial
    app = MyApp()
    # arrage gui
    app.gui_arrang()
    # run tkinter
    tk.mainloop()

if __name__ == '__main__':
    main()
