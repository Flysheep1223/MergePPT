import tkinter as tk
from tkinter import filedialog, messagebox
from spire.presentation import Presentation, FileFormat


class PPTMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT Merger")

        self.file_paths = []

        self.add_button = tk.Button(root, text="Add PPT File", command=self.add_file)
        self.add_button.pack(pady=10)

        self.merge_button = tk.Button(root, text="Merge PPT Files", command=self.merge_files)
        self.merge_button.pack(pady=10)

        self.listbox = tk.Listbox(root, width=50)
        self.listbox.pack(pady=10)

        self.save_path = tk.StringVar()
        self.save_path_entry = tk.Entry(root, textvariable=self.save_path, width=50)
        self.save_path_entry.pack(pady=10)

        self.browse_button = tk.Button(root, text="Browse Save Location", command=self.browse_save_location)
        self.browse_button.pack(pady=10)

    def add_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PPTX files", "*.pptx")])
        if file_path:
            self.file_paths.append(file_path)
            self.listbox.insert(tk.END, file_path)

    def browse_save_location(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PPTX files", "*.pptx")])
        if save_path:
            self.save_path.set(save_path)

    def merge_files(self):
        if not self.file_paths:
            messagebox.showwarning("Warning", "No PPT files selected!")
            return

        if not self.save_path.get():
            messagebox.showwarning("Warning", "No save location specified!")
            return

        try:
            pres1 = Presentation()
            pres1.LoadFromFile(self.file_paths[0])

            for file_path in self.file_paths[1:]:
                pres = Presentation()
                pres.LoadFromFile(file_path)
                for slide in pres.Slides:
                    pres1.Slides.AppendBySlide(slide)
                pres.Dispose()

            pres1.SaveToFile(self.save_path.get(), FileFormat.Pptx2019)
            pres1.Dispose()

            messagebox.showinfo("Success", "PPT files merged successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PPTMergerApp(root)
    root.mainloop()
