from src.gui import GestorReceitasGUI
import tkinter as tk

def main():
    root = tk.Tk()
    app = GestorReceitasGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
