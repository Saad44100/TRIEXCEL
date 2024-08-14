import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser
import os

class AnalyzeTab:
    def __init__(self, parent):
        self.parent = parent
        self.conditions = []
        self.sorting_order = {}
        self.is_blinking = False  # Variable pour contrôler le clignotement

        self.main_tab = ttk.Frame(self.parent.notebook)
        self.parent.notebook.add(self.main_tab, text="Analyser et Rechercher")

        # Texte de bienvenue
        self.welcome_label = tk.Label(self.main_tab, text="Bonjour et bienvenue !\nUtiliser cet onglet pour rechercher avec conditions dans plusieurs fichiers facilement.\n", justify=tk.CENTER, padx=10, pady=10)
        self.welcome_label.pack(anchor='center')

        # Lien cliquable pour le guide
        self.guide_link = tk.Label(self.main_tab, text="Cliquez ici pour accéder au guide d'utilisation (VRAIMENT UTILE !!!) ;)", fg="grey", cursor="hand2", padx=10, pady=10)
        self.guide_link.pack(anchor='center')
        self.guide_link.bind("<Button-1>", self.open_guide)

        # Commencer le clignotement
        self.blink()

        # Bouton pour charger plusieurs fichiers Excel
        self.load_button = tk.Button(self.main_tab, text="Charger les fichiers Excel", command=self.load_files)
        self.load_button.pack(pady=10)

        # Bouton pour concaténer les fichiers chargés
        self.merge_button = tk.Button(self.main_tab, text="Concaténation",
                                      command=self.concatenate_dataframes)
        self.merge_button.pack(pady=5)

        # Cadre pour les conditions dynamiques
        self.conditions_frame = tk.Frame(self.main_tab)
        self.conditions_frame.pack(pady=10)

        # Bouton pour ajouter une condition
        self.add_condition_button = tk.Button(self.main_tab, text="Ajouter une condition", command=self.add_condition)
        self.add_condition_button.pack(pady=5)

        # Bouton pour appliquer les conditions
        self.apply_button = tk.Button(self.main_tab, text="Appliquer les conditions", command=self.apply_conditions)
        self.apply_button.pack(pady=10)

        # Bouton pour exporter les résultats
        self.export_button = tk.Button(self.main_tab, text="Exporter résultats sous Excel", command=self.export_results)
        self.export_button.pack(pady=5)

        # Bouton quitter
        self.quit_button = tk.Button(self.main_tab, text="Quitter", command=self.parent.root.quit)
        self.quit_button.pack(pady=10)

        # Compteur de lignes
        self.line_count_label = tk.Label(self.main_tab, text="")
        self.line_count_label.pack(pady=5, anchor='e')

        # Création du tableau pour afficher les résultats
        self.result_frame = tk.Frame(self.main_tab)
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=20)

        self.result_tree = ttk.Treeview(self.result_frame, show="headings")
        self.result_tree.pack(fill=tk.BOTH, expand=True)

        self.result_scroll_x = ttk.Scrollbar(self.result_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.result_scroll_y = ttk.Scrollbar(self.result_frame, orient="vertical", command=self.result_tree.yview)
        self.result_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.result_tree.configure(xscrollcommand=self.result_scroll_x.set, yscrollcommand=self.result_scroll_y.set)

        # Ajouter des colonnes fictives pour illustrer le problème
        columns = ['Column 1', 'Column 2', 'Column 3', 'Column 4', 'Column 5']
        self.result_tree["columns"] = columns

        # Configurer les colonnes avec des largeurs par défaut et la possibilité de s'étendre
        for col in columns:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=100, minwidth=100, stretch=tk.YES, anchor=tk.W)

        self.result_tree.bind("<Button-1>", self.sort_column)

    def blink(self):
        if self.is_blinking:
            self.guide_link.config(fg="blue")
        else:
            self.guide_link.config(fg="red")  # Changez "red" pour "white" ou "" pour rendre le texte invisible

        self.is_blinking = not self.is_blinking
        # Répéter la fonction toutes les 500 millisecondes (0.5 secondes)
        self.parent.root.after(500, self.blink)

    def open_guide(self, event):
        # Chemin complet du fichier
        guide_path = os.path.join(os.getcwd(), "Guide_SEM.mht")
        # Ouvre le fichier avec l'application par défaut
        webbrowser.open(f"file://{guide_path}")

    def load_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_paths:
            try:
                self.dataframes = []
                for file_path in file_paths:
                    df = pd.read_excel(file_path)
                    self.dataframes.append(df)
                messagebox.showinfo("Succès", f"{len(file_paths)} fichiers chargés avec succès ! ;)")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors du chargement des fichiers: {e} :(")

    def concatenate_dataframes(self):
        if self.dataframes:
            try:
                # Concaténer les dataframes
                self.combined_df = pd.concat(self.dataframes, ignore_index=True)

                self.result_tree.delete(*self.result_tree.get_children())
                self.result_tree["columns"] = list(self.combined_df.columns)
                messagebox.showinfo("Succès", "Concaténation OK !")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la concaténation des fichiers: {e}")
        else:
            messagebox.showwarning("Attention", "Aucun fichier à concaténer. Veuillez charger des fichiers d'abord.")

    def add_condition(self):
        if hasattr(self, 'combined_df') and self.combined_df is not None:
            condition_frame = tk.Frame(self.conditions_frame)
            condition_frame.pack(fill=tk.X, pady=5)

            column_combobox = ttk.Combobox(condition_frame, state="readonly", values=list(self.combined_df.columns))
            column_combobox.pack(side=tk.LEFT, padx=5)

            condition_entry = tk.Entry(condition_frame)
            condition_entry.pack(side=tk.LEFT, padx=5)

            remove_button = tk.Button(condition_frame, text="X", command=lambda: self.remove_condition(condition_frame))
            remove_button.pack(side=tk.LEFT, padx=5)

            self.conditions.append((column_combobox, condition_entry))
        else:
            messagebox.showwarning("Attention", "Veuillez concaténer les fichiers avant d'ajouter des conditions.")

    def remove_condition(self, frame):
        frame.destroy()
        self.conditions = [(col, cond) for col, cond in self.conditions if col.winfo_exists()]

    def apply_conditions(self):
        if hasattr(self, 'combined_df') and self.combined_df is not None:
            try:
                query = self.build_query()
                if query:
                    filtered_df = self.combined_df.query(query)
                    self.display_results(filtered_df)
                else:
                    messagebox.showwarning("Attention", "Veuillez ajouter des conditions.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'application des conditions: {e}")
        else:
            messagebox.showwarning("Attention", "Veuillez concaténer les fichiers avant d'appliquer les conditions.")

    def build_query(self):
        query_parts = []
        for column_combobox, condition_entry in self.conditions:
            column = column_combobox.get()
            condition = condition_entry.get()

            if column and condition:
                # Échapper le nom de la colonne avec des backticks pour gérer les accents et espaces
                column = f'`{column}`'

                # Traiter les conditions spécifiques comme "non vide" et "vide"
                if condition.lower() == "non vide":
                    query_parts.append(f"{column}.notnull()")
                elif condition.lower() == "vide":
                    query_parts.append(f"{column}.isnull()")
                else:
                    # Traiter les conditions numériques ou textuelles en formatant correctement
                    query_parts.append(f"{column} {condition}")

        return " & ".join(query_parts)

    def display_results(self, results_df):
        self.result_tree.delete(*self.result_tree.get_children())
        self.result_tree["columns"] = list(results_df.columns)

        if not results_df.empty:
            for col in results_df.columns:
                self.result_tree.heading(col, text=col)
                self.result_tree.column(col, anchor=tk.W, width=150)

            for _, row in results_df.iterrows():
                values = ["" if pd.isna(item) else item for item in row]
                self.result_tree.insert("", "end", values=values)

            self.line_count_label.config(text=f"Nombre de lignes : {len(results_df)}")
        else:
            messagebox.showinfo("Résultats", "Aucun résultat trouvé pour les conditions données.")
            self.line_count_label.config(text="Nombre de lignes : 0")

    def sort_column(self, event):
        column = self.result_tree.identify_column(event.x)[1:]
        column_name = self.result_tree["columns"][int(column) - 1]
        df = pd.DataFrame([self.result_tree.item(item)["values"] for item in self.result_tree.get_children()],
                          columns=self.result_tree["columns"])

        if self.sorting_order.get(column_name, True):
            df = df.sort_values(by=column_name, ascending=True)
        else:
            df = df.sort_values(by=column_name, ascending=False)

        self.sorting_order[column_name] = not self.sorting_order.get(column_name, True)

        self.result_tree.delete(*self.result_tree.get_children())
        for _, row in df.iterrows():
            self.result_tree.insert("", "end", values=list(row))

    def export_results(self):
        if self.result_tree.get_children():
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                try:
                    rows = [self.result_tree.item(item)["values"] for item in self.result_tree.get_children()]
                    results_df = pd.DataFrame(rows, columns=self.result_tree["columns"])
                    results_df.to_excel(file_path, index=False)
                    messagebox.showinfo("Succès", f"Résultats exportés avec succès dans '{file_path}'")
                except Exception as e:
                    messagebox.showerror("Erreur", f"Erreur lors de l'exportation des résultats: {e}")
        else:
            messagebox.showwarning("Attention", "Aucun résultat à exporter.")
