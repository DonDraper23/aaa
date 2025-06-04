import tkinter as tk
from tkinter import messagebox
import subprocess
import tempfile

# Sample product data for textile fabrics
PRODUCT_DATA = {
    "TX001": {
        "name": "Cotton Twill",
        "composition": "100% cotton",
        "width": "150cm",
        "weight": "180gsm",
        "color": "Blue"
    },
    "TX002": {
        "name": "Polyester Satin",
        "composition": "100% polyester",
        "width": "145cm",
        "weight": "120gsm",
        "color": "Red"
    },
    # Add more product entries as needed
}

def get_product_info(pid):
    """Return formatted product info string for the given product ID."""
    info = PRODUCT_DATA.get(pid.upper())
    if not info:
        return None
    lines = [f"Product ID: {pid.upper()}"]
    for key, value in info.items():
        lines.append(f"{key.capitalize()}: {value}")
    return "\n".join(lines)

def print_info(text):
    """Send the provided text to the default printer on macOS using lpr."""
    with tempfile.NamedTemporaryFile(delete=False, mode="w", suffix=".txt") as tmp:
        tmp.write(text)
        tmp_filename = tmp.name
    try:
        subprocess.run(["lpr", tmp_filename], check=True)
    except Exception as e:
        messagebox.showerror("Print Error", f"Failed to print: {e}")

class ProductApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Textile Product Info")
        self.geometry("400x250")

        self.pid_label = tk.Label(self, text="Enter Product ID:")
        self.pid_label.pack(pady=5)
        self.pid_entry = tk.Entry(self)
        self.pid_entry.pack(fill="x", padx=20)

        self.info_text = tk.Text(self, height=8)
        self.info_text.pack(fill="both", expand=True, padx=20, pady=5)

        self.button_frame = tk.Frame(self)
        self.button_frame.pack(pady=5)

        self.search_button = tk.Button(self.button_frame, text="Search", command=self.search)
        self.search_button.pack(side="left", padx=5)

        self.print_button = tk.Button(self.button_frame, text="Print", command=self.print_current)
        self.print_button.pack(side="left", padx=5)

    def search(self):
        pid = self.pid_entry.get().strip()
        if not pid:
            messagebox.showwarning("Input Error", "Please enter a product ID")
            return
        info = get_product_info(pid)
        self.info_text.delete("1.0", tk.END)
        if info:
            self.info_text.insert(tk.END, info)
        else:
            self.info_text.insert(tk.END, f"No information found for ID: {pid}")

    def print_current(self):
        text = self.info_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("No Data", "No product information to print")
            return
        print_info(text)

if __name__ == "__main__":
    app = ProductApp()
    app.mainloop()
