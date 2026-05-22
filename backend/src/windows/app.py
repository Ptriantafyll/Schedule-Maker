"""
Windows GUI for Schedule Maker
"""

import customtkinter as ctk

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    """
    Main application class for the Windows GUI.
    Inherits from ctk.CTk, which is a custom Tkinter window.
    """

    def __init__(self):
        super().__init__()  # Initialize the underlying CTk window

        self.title("My Windows App")

        window_width = 400
        window_height = 240
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int((screen_width / 2) - (window_width / 2))
        center_y = int((screen_height / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

        self.label = ctk.CTkLabel(
            self, text="Welcome to my app!", font=("Arial", 18))
        self.label.pack(pady=40)

        self.button = ctk.CTkButton(
            self, text="Click Me", command=self.button_click)
        self.button.pack(pady=10)

    def button_click(self):
        """Event handler for button click. Updates the label text."""
        self.label.configure(text="Button was clicked!")


if __name__ == "__main__":
    app = App()
    app.mainloop()
