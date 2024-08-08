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

        # Create a frame for the listbox and scrollbars
        self.frame = tk.Frame(root)
        self.frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # Create a horizontal scrollbar
        self.h_scrollbar = tk.Scrollbar(self.frame, orient=tk.HORIZONTAL)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Create a vertical scrollbar
        self.v_scrollbar = tk.Scrollbar(self.frame)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create the listbox
        self.listbox = tk.Listbox(self.frame, width=50, xscrollcommand=self.h_scrollbar.set, yscrollcommand=self.v_scrollbar.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure the scrollbars
        self.h_scrollbar.config(command=self.listbox.xview)
        self.v_scrollbar.config(command=self.listbox.yview)

        # Add move up and move down buttons
        self.move_up_button = tk.Button(root, text="Move Up", command=self.move_up)
        self.move_up_button.pack(pady=5)

        self.move_down_button = tk.Button(root, text="Move Down", command=self.move_down)
        self.move_down_button.pack(pady=5)

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

    def move_up(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            return
        for i in selected_indices:
            if i == 0:
                continue
            self.file_paths[i - 1], self.file_paths[i] = self.file_paths[i], self.file_paths[i - 1]
            self.listbox.delete(i)
            self.listbox.insert(i - 1, self.file_paths[i - 1])
            self.listbox.selection_set(i - 1)

    def move_down(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            return
        for i in reversed(selected_indices):
            if i == self.listbox.size() - 1:
                continue
            self.file_paths[i + 1], self.file_paths[i] = self.file_paths[i], self.file_paths[i + 1]
            self.listbox.delete(i)
            self.listbox.insert(i + 1, self.file_paths[i + 1])
            self.listbox.selection_set(i + 1)

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
    root.geometry("500x500")
    app = PPTMergerApp(root)
    root.mainloop()
