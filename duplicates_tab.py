import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class DuplicatesTab:
    def __init__(self, parent):
        self.parent = parent
        self.sorting_order = {}

        self.duplicates_tab = ttk.Frame(self.parent.notebook)
        self.parent.notebook.add(self.duplicates_tab, text="Détection de Doublons")

        # Bouton pour charger plusieurs fichiers Excel
        self.load_button_duplicates = tk.Button(self.duplicates_tab, text="Charger les fichiers Excel",
                                                command=self.load_files)
        self.load_button_duplicates.pack(pady=10)

        # Bouton pour concaténer les fichiers chargés
        self.merge_button_duplicates = tk.Button(self.duplicates_tab, text="Concaténation",
                                                 command=self.concatenate_dataframes)
        self.merge_button_duplicates.pack(pady=5)

        # Bouton pour afficher les doublons
        self.duplicate_button = tk.Button(self.duplicates_tab, text="Afficher les doublons",
                                          command=self.show_duplicates)
        self.duplicate_button.pack(pady=10)

        # Bouton pour exporter les résultats
        self.export_button_duplicates = tk.Button(self.duplicates_tab, text="Exporter les résultats",
                                                  command=self.export_duplicates_results)
        self.export_button_duplicates.pack(pady=5)

        # Bouton quitter
        self.quit_button = tk.Button(self.duplicates_tab, text="Quitter", command=self.parent.root.quit)
        self.quit_button.pack(pady=10)

        # Compteur de lignes pour les doublons
        self.line_count_label_duplicates = tk.Label(self.duplicates_tab, text="")
        self.line_count_label_duplicates.pack(pady=5, anchor='e')

        # Cadre pour afficher les résultats de doublons
        self.result_frame_duplicates = tk.Frame(self.duplicates_tab)
        self.result_frame_duplicates.pack(fill=tk.BOTH, expand=True, pady=20)

        self.result_tree_duplicates = ttk.Treeview(self.result_frame_duplicates, show="headings")
        self.result_tree_duplicates.pack(fill=tk.BOTH, expand=True)

        self.result_scroll_x_duplicates = ttk.Scrollbar(self.result_frame_duplicates, orient="horizontal", command=self.result_tree_duplicates.xview)
        self.result_scroll_x_duplicates.pack(side=tk.BOTTOM, fill=tk.X)

        self.result_scroll_y_duplicates = ttk.Scrollbar(self.result_frame_duplicates, orient="vertical", command=self.result_tree_duplicates.yview)
        self.result_scroll_y_duplicates.pack(side=tk.RIGHT, fill=tk.Y)

        self.result_tree_duplicates.configure(xscrollcommand=self.result_scroll_x_duplicates.set, yscrollcommand=self.result_scroll_y_duplicates.set)

        # Ajouter des colonnes fictives pour illustrer le problème (à remplacer par les colonnes réelles)
        columns = ['Column A', 'Column B', 'Column C', 'Column D', 'Column E']
        self.result_tree_duplicates["columns"] = columns

        # Configurer les colonnes avec des largeurs par défaut et la possibilité de s'étendre
        for col in columns:
            self.result_tree_duplicates.heading(col, text=col)
            self.result_tree_duplicates.column(col, width=100, minwidth=100, stretch=tk.YES, anchor=tk.W)

        self.result_tree_duplicates.bind("<Button-1>", self.sort_column_duplicates)

    def load_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_paths:
            try:
                self.dataframes = []
                for file_path in file_paths:
                    df = pd.read_excel(file_path)
                    self.dataframes.append(df)
                messagebox.showinfo("Succès", f"{len(file_paths)} fichiers chargés avec succès!")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors du chargement des fichiers: {e}")

    def concatenate_dataframes(self):
        if self.dataframes:
            try:
                # Concaténer les dataframes
                self.combined_df = pd.concat(self.dataframes, ignore_index=True)

                self.result_tree_duplicates.delete(*self.result_tree_duplicates.get_children())
                self.result_tree_duplicates["columns"] = list(self.combined_df.columns)
                messagebox.showinfo("Succès", "Concaténation des fichiers effectuée avec succès!")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la concaténation des fichiers: {e}")
        else:
            messagebox.showwarning("Attention", "Aucun fichier à concaténer. Veuillez charger des fichiers d'abord.")

    def show_duplicates(self):
        if len(self.dataframes) < 2:
            messagebox.showwarning("Attention", "Veuillez charger au moins deux fichiers pour rechercher des doublons.")
            return

        if hasattr(self, 'combined_df') and self.combined_df is not None:
            duplicate_window = tk.Toplevel(self.parent.root)
            duplicate_window.title("Afficher les doublons")

            tk.Label(duplicate_window, text="Sélectionnez la colonne à vérifier pour les doublons:").pack(pady=5)
            column_combobox = ttk.Combobox(duplicate_window, state="readonly", values=list(self.combined_df.columns))
            column_combobox.pack(pady=5)

            tk.Button(duplicate_window, text="Afficher les doublons",
                      command=lambda: self.display_duplicates(column_combobox.get(), duplicate_window)).pack(pady=10)
        else:
            messagebox.showwarning("Attention", "Veuillez concaténer les fichiers avant d'afficher les doublons.")

    def display_duplicates(self, column, window):
        if column:
            try:
                # Trouver les valeurs communes entre tous les fichiers pour une colonne spécifique
                common_values = pd.Series(self.dataframes[0][column])
                for df in self.dataframes[1:]:
                    common_values = common_values[common_values.isin(df[column])]

                # Filtrer les dataframes d'origine pour ne garder que les lignes contenant les valeurs communes
                duplicate_df = pd.concat([df[df[column].isin(common_values)] for df in self.dataframes])

                self.result_tree_duplicates.delete(*self.result_tree_duplicates.get_children())
                self.result_tree_duplicates["columns"] = list(duplicate_df.columns)

                if not duplicate_df.empty:
                    for col in duplicate_df.columns:
                        self.result_tree_duplicates.heading(col, text=col)
                        self.result_tree_duplicates.column(col, anchor=tk.W, width=150)

                    for _, row in duplicate_df.iterrows():
                        values = ["" if pd.isna(item) else item for item in row]
                        self.result_tree_duplicates.insert("", "end", values=values)

                    self.line_count_label_duplicates.config(text=f"Nombre de lignes : {len(duplicate_df)}")
                else:
                    messagebox.showinfo("Résultats", "Aucun doublon trouvé.")
                    self.line_count_label_duplicates.config(text="Nombre de lignes : 0")

                window.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'affichage des doublons: {e}")
        else:
            messagebox.showwarning("Attention", "Veuillez sélectionner une colonne.")

    def sort_column_duplicates(self, event):
        column = self.result_tree_duplicates.identify_column(event.x)[1:]
        column_name = self.result_tree_duplicates["columns"][int(column) - 1]
        df = pd.DataFrame(
            [self.result_tree_duplicates.item(item)["values"] for item in self.result_tree_duplicates.get_children()],
            columns=self.result_tree_duplicates["columns"])

        if self.sorting_order.get(column_name, True):
            df = df.sort_values(by=column_name, ascending=True)
        else:
            df = df.sort_values(by=column_name, ascending=False)

        self.sorting_order[column_name] = not self.sorting_order.get(column_name, True)

        self.result_tree_duplicates.delete(*self.result_tree_duplicates.get_children())
        for _, row in df.iterrows():
            self.result_tree_duplicates.insert("", "end", values=list(row))

    def export_duplicates_results(self):
        if self.result_tree_duplicates.get_children():
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                try:
                    rows = [self.result_tree_duplicates.item(item)["values"] for item in
                            self.result_tree_duplicates.get_children()]
                    results_df = pd.DataFrame(rows, columns=self.result_tree_duplicates["columns"])
                    results_df.to_excel(file_path, index=False)
                    messagebox.showinfo("Succès", f"Résultats exportés avec succès dans '{file_path}'")
                except Exception as e:
                    messagebox.showerror("Erreur", f"Erreur lors de l'exportation des résultats: {e}")
        else:
            messagebox.showwarning("Attention", "Aucun résultat à exporter.")
