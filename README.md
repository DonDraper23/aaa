# Textile Product Info Manager

This repository contains a small GUI application for managing textile fabric information. It allows you to enter a product ID and view the associated details. You can also print the displayed information using your macOS default printer.

## Requirements

- Python 3 (tested with Python 3.8+)
- Tkinter (comes with the standard Python installation on macOS)

## Usage

1. Open a terminal and navigate to this repository directory.
2. Run the application:

```bash
python3 main.py
```

3. Enter a product ID (e.g., `TX001` or `TX002`) and click **Search** to view the product details.
4. Click **Print** to send the displayed information to your default printer (uses the `lpr` command on macOS).

You can extend the `PRODUCT_DATA` dictionary in `main.py` to include additional fabric products.

## Packaging as a macOS App

To create a standalone `.app` bundle using **py2app**:

1. Install py2app:

   ```bash
   pip install py2app
   ```

2. Run the build script:

   ```bash
   python3 setup.py py2app
   ```

3. After the build completes, the application bundle will be located in the `dist` directory. You can move the generated `.app` to your Applications folder.

The provided `setup.py` is configured with basic options and uses the GUI script `main.py` as the entry point.
