import tkinter as tk
from tkinter import messagebox
from resightingsToExcel import get_resightings_to_Excel

class ResightingsGUI(tk.Tk):

    info_text = ("This tool will create an Excel sheet from all the submitted resightings "
                "from the time range specified by the start and end date.\n It will also "
                "save copies of each full size image and the associated form data in "
                "their own folders for additional viewing.\n"
                "Enter the dates in month-day-year format.")

    instr_text = "Enter the dates in month-day-year format."

    err_title = "Invalid Date Format"

    err_msg = "Please enter dates in day-month-year format."
    

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("Resightings Export Tool")
        frame = tk.Frame(self)

        # Information label
        info_label = tk.Label(frame, text=self.info_text)
        instr_label = tk.Label(frame, text=self.instr_text, font="-weight bold")

        # Make date labels
        date_label1 = tk.Label(frame, text="Start Date", width=40)
        date_label2 = tk.Label(frame, text="End Date", width=40)

        # Make text entry boxes. Instance variables to get input later
        self.date_entry1 = tk.Entry(frame)
        self.date_entry2 = tk.Entry(frame)

        # Make run button, hook up to resightingsToExcel function
        run_button = tk.Button(frame, text="Run", command=self._create_sheet_with_dates)

        # Pack all in order
        items = [frame, 
                 info_label, 
                 instr_label, 
                 date_label1, 
                 self.date_entry1,
                 date_label2, 
                 self.date_entry2, 
                 run_button]
        for item in items:
            item.pack()


    def _create_sheet_with_dates(self):
        start = self.date_entry1.get().strip()
        end = self.date_entry2.get().strip()
        dest = "Resightings_from_{d1}_to_{d2}.xlsx".format(d1=start, d2=end)
        try:
            get_resightings_to_Excel(start=start, end=end, dest=dest)
        except RuntimeError:
            messagebox.showerror(self.err_title, self.err_msg)


if __name__ == "__main__":
    gui = ResightingsGUI()
    gui.mainloop()
