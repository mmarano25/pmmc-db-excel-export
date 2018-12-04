import tkinter as tk

if __name__ == "__main__":
    master = tk.Tk()
    master.title("Resightings Export Tool")
    frame = tk.Frame(master)

    # Information label
    info_text = ("This tool will create an Excel sheet from all the submitted resightings "
                  "from the time range specified by the start and end date.\n It will also "
                  "save copies of each full size image and the associated form data in "
                  "their own folders for additional viewing.\n"
                  "Enter the dates in month/day/year format.")
    info_label = tk.Label(frame, text=info_text)
    instr_text = "Enter the dates in month/day/year format."
    instr_label = tk.Label(frame, text=instr_text, font="-weight bold")

    # Make date labels
    date_label1 = tk.Label(frame, text="Start Date", width=40)
    date_label2 = tk.Label(frame, text="End Date", width=40)

    # Make text entry boxes
    e1 = tk.Entry(frame)
    e2 = tk.Entry(frame)

    # Make run button
    run_button = tk.Button(frame, text="Run")

    # Pack all in order
    frame.pack()
    info_label.pack()
    instr_label.pack()
    date_label1.pack()
    e1.pack()
    date_label2.pack()
    e2.pack()
    run_button.pack()

    tk.mainloop()

