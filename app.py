import tkinter
import tkinter.messagebox
import customtkinter
from tkinter import filedialog
from tkinter import messagebox
import os
import openpyxl
import pandas as pd
from datetime import datetime

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"



# Cria um novo workbook e seleciona a planilha ativa

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("Desenvolvido por Jhonatan")
        self.geometry(f"{600}x{250}")
        self.maxsize(width=600, height=250)
        self.minsize(width=600, height=250)

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # create main entry and button
        self.entry = customtkinter.CTkLabel(self, text='', bg_color='light blue', anchor="w", justify="left", padx=10, pady=10, text_color='black')
        self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")

        def chosseFile():
            self.entry.configure(text='')
            file = filedialog.askopenfilename(
                                initialdir="/",
                                title="Selecione um arquivo CSV",
                                filetypes=(("CSV", "*.csv"),)
                            )
            df = pd.read_csv(file, encoding='ISO-8859-1', delimiter=';')

            if file:
                file_name = os.path.basename(file)
          
    
            filename = file
            print(str(file))
            with open(str(file), 'r') as f:
                filename = f.name

            extension = os.path.splitext(filename)[1]  # Obter a extensão do arquivo

            if extension == ".csv":

                def loopTable(event):

                    while len(entryNameTable.get()) < 2:
                        return messagebox.showwarning(title='Erro', message='O nome da planilha deve conter mais de 03 caracteres')

                    while str(entryMonth.get())[0] == ' ' or len(entryMonth.get()) > 2 or len(entryMonth.get()) < 2 or entryMonth.get().isalpha() or str(entryMonth.get())[0].isalpha() or str(entryMonth.get())[1].isalpha():
                        return messagebox.showwarning(title='Erro', message='Por favor! Informe o mês no formato XX. Sem espaço apenas números')
                    
                    global new_workbook
                    new_workbook = openpyxl.Workbook()
                    new_worksheet = new_workbook.active

                    # Define o cabeçalho da tabela
                    header = ['Nome do Hóspede', 'Data de Nascimento', 'Email do Hóspede', 'Telefone']

                    # Adiciona o cabeçalho à primeira linha da planilha
                    new_worksheet.append(header)

                    
                    for index, birthDay in enumerate(df['Data de Nascimento']): 
                        birthDay = str(birthDay)
                        name = df.loc[index, 'Nome do Hóspede']
                        tel = df.loc[index, 'Telefone']
                        email = df.loc[index, 'Email do Hóspede']
                        
                        if birthDay[3:5] == str(entryMonth.get()):
                            new_worksheet.append([name, birthDay, email, tel])

                    
                    
                    
                    agora = datetime.now().strftime('%Y-%m-%d')
                    nome_arquivo = f'{str(entryNameTable.get())}_{agora}.xlsx'
                    new_workbook.save(nome_arquivo)
                    entryNameTable.delete(0, 'end')
                    entryMonth.delete(0, 'end')
                            
                            

                    

                def chooseFileWindow():
                    
                    self.entry.configure(text=f'{file_name}')
                    windowsBirth = customtkinter.CTkToplevel()
                    windowsBirth.geometry('400x220')
                    windowsBirth.maxsize(width=400, height=220)
                    windowsBirth.minsize(width=400, height=220)
                    windowsBirth.title('Adicione as informções')
                    windowsBirth.focus_force()
                    windowsBirth.grab_set()
                    windowsBirth.lift()


                    global entryMonth

                    labelFileName = customtkinter.CTkLabel(windowsBirth, text='Informe o nome do arquivo. Ex: Dezembro', width=350, anchor="w", justify="left")
                    labelFileName.grid( padx=(20, 0), pady=(5, 5), sticky="nsew")

                    global entryNameTable
                    entryNameTable = customtkinter.CTkEntry(windowsBirth, placeholder_text="Informe o nome do arquivo", width=350)
                    entryNameTable.grid( padx=(20, 0), pady=(10, 10), sticky="nsew")

                    labelMonthBirth = customtkinter.CTkLabel(windowsBirth, text='Informe o mês do aniversário. Ex: 02', width=350, anchor="w", justify="left")
                    labelMonthBirth.grid( padx=(20, 0), pady=(5, 5), sticky="nsew")


                    
                    entryMonth = customtkinter.CTkEntry(windowsBirth, placeholder_text="Informe o mês no formato XX", width=350, corner_radius=2)
                    entryMonth.grid( padx=(20, 0), pady=(10, 10), sticky="nsew")

                    

          

                    entryMonth.bind('<Return>', loopTable)
                    entryNameTable.bind('<Return>', loopTable)

                chooseFileWindow()

                    
            else:
                file = ''
                messagebox.showwarning(title='Arquivo incompatível', message='Selecione um arquivo no formato csv')
                return chosseFile()
            
                      
        
            
            
        self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text='Chosse file',text_color=("gray10", "#DCE4EE"), command=chosseFile)
        self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")


       

    def open_input_dialog_event(self):
        dialog = customtkinter.CTkInputDialog(text="Type in a number:", title="CTkInputDialog")
        print("CTkInputDialog:", dialog.get_input())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def sidebar_button_event(self):
        print("sidebar_button click")


if __name__ == "__main__":
    app = App()
    app.mainloop()