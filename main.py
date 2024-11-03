import tkinter as tk
from gui import YahooAuctionScraperGUI

def main():
    root = tk.Tk()
    app = YahooAuctionScraperGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()