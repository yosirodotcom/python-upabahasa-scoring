import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import pandas as pd

rumus_grammar = pd.DataFrame({
    "g1": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40],
    "g2": [21, 22, 24, 25, 26, 28, 29, 30, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 57, 58, 59, 61, 63, 64, 66, 67]
})

rumus_listening = pd.DataFrame({
    "l1": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50],
    "l2": [25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 36, 37, 38, 39, 40, 41, 41, 42, 43, 43, 44, 44, 45, 46, 46, 47, 48, 48, 49, 49, 50, 50, 51, 52, 52, 53, 54, 55, 56, 56, 57, 58, 59, 60, 61, 62, 63, 64, 66, 68]
})

rumus_reading = pd.DataFrame({
    "r1": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50],
    "r2": [22, 23, 24, 25, 26, 27, 28, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 41, 42, 43, 44, 45, 45, 46, 47, 47, 48, 48, 49, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 63, 65, 66, 67]
})



def select_files():
    filenames = filedialog.askopenfilenames(title="Select Excel Files", filetypes=[("Excel Files", "*.xlsx")])
    file_paths = []
    for filename in filenames:
        file_paths.append(filename)
    for widget in root.pack_slaves():
        if isinstance(widget, tk.Label):
            widget.destroy()
    create_table(file_paths)
    return file_paths
def create_table(file_paths):
    df = pd.concat([pd.read_excel(file_path) for file_path in file_paths])
    df = df[["No. Peserta", "Section 1", "Section 2", "Section 3"]]
    df.columns = ["ID", "l1", "g1", "r1"]
    df = df.merge(rumus_listening, on="l1", how="left")
    df = df.merge(rumus_grammar, on="g1", how="left")
    df = df.merge(rumus_reading, on="r1", how="left")
    df["skor"] = round((df["l2"]+df["g2"]+df["r2"])*10/3)
    df["hadir"] = "TRUE"
    df = df.sort_values(by=['ID']).reset_index(drop=True)
    print(df.describe())
# Create the main window
root = ctk.CTk()
root.title("Select Multiple Excel Files")
root.geometry("300x300")  # Set the window size to 300x300 pixels
# Create the button to select files
button = ctk.CTkButton(root, text="Select Files", command=select_files)
button.pack(pady=50)  # Set the button's y position to 50 pixels
if __name__ == "__main__":
    # Run the main loop
    root.mainloop()