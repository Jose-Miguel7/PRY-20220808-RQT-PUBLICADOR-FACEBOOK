from pathlib import Path
from tkinter import Tk, Canvas, Entry, Button, PhotoImage, filedialog, messagebox, StringVar
from publish import BotFacebookMarketplace
from manage_excel import update_image_excel, create_excel_format

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


class App(Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.resizable(False, False)
        self.geometry("1000x650")
        self.configure(bg="#FFFFFF")
        self.title("Publicador Facebook Marketplace")

        self.excel = None
        self.update_excel = None
        self.directory = None

        self.directory_images = StringVar()
        self.directory_excel_images = StringVar()
        self.directory_excel_publish = StringVar()
        self.email = StringVar()
        self.password = StringVar()

        canvas = Canvas(self, bg="#FFFFFF", height=650, width=1000, bd=0, highlightthickness=0, relief="ridge")

        canvas.place(x=0, y=0)
        canvas.create_rectangle(602.0, 0.0, 1000.0, 720.0, fill="#D9D9D9", outline="")

        self.entry_image_2 = PhotoImage(file=relative_to_assets("entry_2.png"))
        canvas.create_image(801.0, 156.5, image=self.entry_image_2)
        email_entry = Entry(bd=0, bg="#F1F1F1", highlightthickness=0, textvariable=self.email)
        email_entry.place(x=647.5, y=139.0, width=307.0, height=33.0)

        self.entry_image_1 = PhotoImage(file=relative_to_assets("entry_1.png"))
        canvas.create_image(801.0, 233.5, image=self.entry_image_1)
        password_entry = Entry(bd=0, bg="#F1F1F1", highlightthickness=0, textvariable=self.password, show="*")
        password_entry.place(x=647.5, y=216.0, width=307.0, height=33.0)

        canvas.create_text(23.0, 37.0, anchor="nw", text="Publicador Facebook Marketplace", fill="#1A73E4",
                           font=("MontserratRoman Regular", 27 * -1, "bold"))

        self.entry_image_3 = PhotoImage(file=relative_to_assets("entry_3.png"))
        canvas.create_image(801.0, 401.0, image=self.entry_image_3)
        excel_publish_entry = Entry(bd=0, bg="#F1F1F1", highlightthickness=0, textvariable=self.directory_excel_publish,
                                    state="disabled")
        excel_publish_entry.place(x=647.0, y=384.0, width=308.0, height=32.0)

        canvas.create_text(630.0, 303.0, anchor="nw", text="Seleccionar Excel con datos", fill="#000000",
                           font=("MontserratRoman SemiBold", 12 * -1, "bold"))

        canvas.create_text(23.0, 180.0, anchor="nw", text="Formato Excel", fill="#000000",
                           font=("MontserratRoman SemiBold", 14 * -1, "bold"))

        canvas.create_text(27.0, 287.0, anchor="nw", text="Seleccionar carpeta que contiene imagenes de los productos",
                           fill="#000000", font=("MontserratRoman SemiBold", 12 * -1, "bold"))

        canvas.create_text(27.0, 324.0, anchor="nw",
                           text="El nombre de cada subcarpeta debe ser el código del producto respectivo. Ej: RQT-102",
                           fill="#000000", font=("MontserratRoman SemiBold", 12 * -1))

        canvas.create_text(27.0, 430.0, anchor="nw", text="Seleccionar Excel con los códigos de los productos",
                           fill="#000000", font=("MontserratRoman SemiBold", 12 * -1))

        self.button_image_1 = PhotoImage(file=relative_to_assets("button_1.png"))
        button_excel_publish = Button(image=self.button_image_1, borderwidth=0, highlightthickness=0,
                                      command=self.search_excel_publish, relief="flat")
        button_excel_publish.place(x=630.0, y=338.0, width=342.0, height=34.0)

        self.entry_image_4 = PhotoImage(file=relative_to_assets("entry_4.png"))
        canvas.create_image(377.0, 378.0, image=self.entry_image_4)
        directory_images_entry = Entry(bd=0, bg="#F1F1F1", highlightthickness=0, textvariable=self.directory_images,
                                       state="disabled")
        directory_images_entry.place(x=214.0, y=361.0, width=326.0, height=32.0)

        self.button_image_2 = PhotoImage(file=relative_to_assets("button_2.png"))
        select_directory_images = Button(image=self.button_image_2, borderwidth=0, highlightthickness=0,
                                         command=self.search_directory_images, relief="flat")
        select_directory_images.place(x=23.0, y=363.0, width=158.0, height=29.0)

        self.entry_image_5 = PhotoImage(file=relative_to_assets("entry_5.png"))
        canvas.create_image(377.0, 475.0, image=self.entry_image_5)
        directory_excel_images_entry = Entry(bd=0, bg="#F1F1F1", highlightthickness=0,
                                             textvariable=self.directory_excel_images, state="disabled")
        directory_excel_images_entry.place(x=214.0, y=458.0, width=326.0, height=32.0)

        self.button_image_3 = PhotoImage(file=relative_to_assets("button_3.png"))
        select_excel_images = Button(image=self.button_image_3, borderwidth=0, highlightthickness=0,
                                     command=self.search_excel_for_image, relief="flat")
        select_excel_images.place(x=23.0, y=460.0, width=158.0, height=29.0)

        self.button_image_4 = PhotoImage(file=relative_to_assets("button_4.png"))
        create_excel_format_button = Button(image=self.button_image_4, borderwidth=0, highlightthickness=0,
                                            command=create_excel_format, relief="flat")
        create_excel_format_button.place(x=23.0, y=228.0, width=530.0, height=29.0)

        self.button_image_5 = PhotoImage(file=relative_to_assets("button_5.png"))
        publish_button = Button(image=self.button_image_5, borderwidth=0, highlightthickness=0,
                                command=self.publish_products, relief="flat")
        publish_button.place(x=630.0, y=498.0, width=342.0, height=42.0)

        canvas.create_text(630.0, 189.0, anchor="nw", text="Contraseña", fill="#000000",
                           font=("MontserratRoman Medium", 12 * -1))

        canvas.create_text(630.0, 109.0, anchor="nw", text="Email", fill="#000000",
                           font=("MontserratRoman Medium", 12 * -1))

    def search_directory_images(self):
        try:
            self.directory = filedialog.askdirectory()
            self.directory_images.set(self.directory)
        except AttributeError:
            pass

    def search_excel_for_image(self):
        try:
            if self.directory:
                self.update_excel = filedialog.askopenfile(filetypes=[("Excel files", "*.xlsx")]).name
                self.directory_excel_images.set(self.update_excel)
                update_image_excel(self.update_excel, self.directory)
            else:
                messagebox.showinfo('Info', 'Falta seleccionar la carpeta con las imagenes')
        except AttributeError:
            pass

    def search_excel_publish(self):
        try:
            self.excel = filedialog.askopenfile(filetypes=[("Excel files", "*.xlsx")]).name
            self.directory_excel_publish.set(self.excel)
        except AttributeError:
            pass

    def publish_products(self):
        if self.excel:
            if self.email.get() and self.password.get():
                BotFacebookMarketplace(self.excel, self.email.get(), self.password.get())
                messagebox.showinfo('Info', 'Proceso terminado')
            else:
                messagebox.showinfo('Info', 'Falta ingresar las credenciales de acceso')
        else:
            messagebox.showinfo('Info', 'Falta ingresar el excel con los datos')


if __name__ == '__main__':
    window = App()
    window.mainloop()
