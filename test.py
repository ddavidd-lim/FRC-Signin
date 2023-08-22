import tkinter as tk

window = tk.Tk()

frames = []  # Create a list to store frames

for i in range(2):
    row_frames = []  # Create a list for frames in each row
    for j in range(2):
        frame = tk.Frame(
            master=window,
            relief=tk.RAISED,
            borderwidth=1
        )
        frame.grid(row=i, column=j)
        label = tk.Label(master=frame, text=f"Row {i}\nColumn {j}")
        label.pack()
        row_frames.append(frame)  # Append the frame to the row list
    frames.append(row_frames)  # Append the row list to the main list

# Configure rows and columns to expand
for i in range(3):
    window.grid_rowconfigure(i, weight=1)
    window.grid_columnconfigure(i, weight=1)

window.mainloop()