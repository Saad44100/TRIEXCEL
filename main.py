import tkinter as tk
from tkinter import ttk
from analyze_tab import AnalyzeTab
from duplicates_tab import DuplicatesTab

class ExcelAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Analyse Excel - SEM BUSINESS COMPANY LTD")

        # Maximiser la fenêtre au démarrage
        self.root.state("zoomed")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)

        # Initialiser les onglets
        self.analyze_tab = AnalyzeTab(self)
        self.duplicates_tab = DuplicatesTab(self)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelAnalyzerApp(root)
    root.mainloop()
