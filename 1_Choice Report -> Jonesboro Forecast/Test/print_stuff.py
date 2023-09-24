import tkinter as tk
from tkinter import font

def close_window():
    root.destroy()

# Create the main window
root = tk.Tk()

# Set the window title
root.title("Results from data transfer")
custom_font = font.Font(family="Arial", size=18)

root.configure(bg="#ECECEC")

# Create a Text widget for displaying text
text_widget = tk.Text(root, font = custom_font)
text_widget.pack()

# Adding text based on changed values
text_widget.insert(tk.END, "NOTE: This text window is temporary! Info from it will be lost when window is closed.lsllslslslslslslslslslslslslslsls\n\n")
text_widget.insert(tk.END, "Neener!\n")

# Create a Close button to close the window
close_button = tk.Button(root, text="Close", command=root.destroy)
close_button.pack()

# Run the main event loop for the results 
root.mainloop()







# import tkinter as tk
# from tkinter import font

# # Create the main window
# root = tk.Tk()

# # Set the window title
# root.title("Spiced Up Window")



# # Configure the window background color
# root.configure(bg="#ECECEC")

# # Create a custom font
# custom_font = font.Font(family="Arial", size=18, weight="bold")

# # Create a Label with custom font
# label = tk.Label(root, text="Spiced Up Window!", font=custom_font, bg="#ECECEC")
# label.pack(pady=20)

# # Run the main event loop
# root.mainloop()