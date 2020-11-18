from tkinter import *

import datetime
import pandas
import pandas as pd
import xlwings as xw

# test1


E = False
T = False
D = False
A = False

# Runs main program
def program():
    # gets directory from text file to save tickets to.
    savelocation = open("SaveLocation.txt")
    sl = savelocation.read()
    print(sl)

    # gets phone list from excel
    def open_phone_list():
        pl = xw.Book("Operator phone numbers.xlsx")

    def get_and_increment_ticket_number(filename="Ticket_Numbers.dat"):
        with open(filename, "a+") as f:
            f.seek(0)
            val0 = int(f.read() or 0) + 1
            f.seek(0)
            f.truncate()
            f.write(str(val0))
            return val0 - 1

    ticket_counter = get_and_increment_ticket_number()
    print("Ticket number {}".format(ticket_counter))

# opens new ticket window
    def new_window():
        # Toplevel object which will
        # be treated as a new window
        program()

    # creates root widget
    root = Tk()
    root.title("Don Blatkewicz Steamer Service")
    root.geometry("1000x800")

    # get arc phone list
    # phone_list = pandas.read_excel("Operator phone numbers.xlsx")
    # get todays date
    today = datetime.date.today()

    # gets work dictionary
    work_dic = {}
    with open("Work Dictionary.txt") as WD:
        for line in WD:
            key, val = line.strip().split(':')
            work_dic[key] = val
    work_keys = list(work_dic.keys())

    driver_dic = {}
    with open("Driver List.txt") as DD:
        for line in DD:
            key, driver_name = line.strip().split(':')
            driver_dic[key] = driver_name
    driver_keys = list(driver_dic.keys())

    # build master frame

    master_frame = LabelFrame(root, text="Don Blatkewicz Steamer Service")
    master_frame.grid(row=0, column=0, padx=10)
    # build billing frame
    billing_frame = LabelFrame(master_frame, text="Billing Information")
    billing_frame.grid(row=0, column=0)
    # build work performed frame
    work_performed_frame = LabelFrame(master_frame, text="Work Performed")
    work_performed_frame.grid(row=1, column=0)
    # build summary frame
    summary_frame = LabelFrame(master_frame, text="Summary")
    summary_frame.grid(row=2, column=0)
    # build driver frame
    driver_frame = LabelFrame(master_frame, text="Driver")
    driver_frame.grid(row=3, column=0)
    # build new ticket button
    new_window = Button(master_frame, text="New Ticket", command=new_window)
    new_window.grid(row=0, column=1)
# build button for contact list
    contact_list = Button(billing_frame, text="Arc Phone list", command=open_phone_list)
    contact_list.grid(row=3, column=4)

    # guts for billing information
    class BillingInformationClass(LabelFrame):
        def __init__(self,):
            # build date label
            date_label = Label(billing_frame, text="Date:")
            date_label.grid(row=0, column=0)
            # build date entry field
            date_entry = Entry(billing_frame, text="Date Entry")
            date_entry.grid(row=0, column=1)
            # inserts current date to entry field
            date_entry.insert(0, today)

            # build charge to label
            charge_to_label = Label(billing_frame, text="Charge to:")
            charge_to_label.grid(row=0, column=2)
            # build charge to entry field
            charge_to_entry = Entry(billing_frame, text="Billing")
            charge_to_entry.grid(row=0, column=3)
            # build work location label
            work_location_label = Label(billing_frame, text="Location:")
            work_location_label.grid(row=1, column=0)
            # build work location entry field
            work_location_entry = Entry(billing_frame, text="Work Location")
            work_location_entry.grid(row=1, column=1)
            # build contact name label
            contact_name_label = Label(billing_frame, text="Contact:")
            contact_name_label.grid(row=2, column=0)
            # build contact name entry field
            contact_name_entry = Entry(billing_frame, text="Contact Phone")
            contact_name_entry.grid(row=2, column=1)
            # build contact phone label
            contact_phone_label = Label(billing_frame, text="Ph #:")
            contact_phone_label.grid(row=2, column=2)
            # build contact phone entry field
            contact_phone_entry = Entry(billing_frame, text="Phone entry")
            contact_phone_entry.grid(row=2, column=3)
            # build ticket number label
            ticket_number = Label(billing_frame, text="Ticket Number:{}".format(ticket_counter))
            ticket_number.grid(row=0, column=4)

            # gets operator phone number if name is in Arc phone list
            def get_operator_phone(name):
                global number
                global phone
                # checks for name in contact list
                phone_list = pandas.read_excel("Operator phone numbers.xlsx")
                phone_list = phone_list.set_index("Name")
                number_list = phone_list.iloc[:, 0]
                if name in number_list:
                    number = phone_list.loc[name, "Work Cellular"]
                    number = str(number)
                    phone = number
                    contact_phone_entry.delete(0, 20)
                    contact_phone_entry.insert(0, phone)

            # collects all billing information and stores it in list.
            def collect_billing_info():
                global billing_information_list
                global charge_to
                global E
                global T
                global D
                global A
                global phone
                charge_to = ""
                billing_information_list = []
                charge_to = charge_to_entry.get()
                work_location = work_location_entry.get()
                contact_name = contact_name_entry.get()
                phone = contact_phone_entry.get()
                phone = phone.replace(" ", "")
                date = date_entry.get()
                get_operator_phone(contact_name)
                billing_information_list = [charge_to, work_location, contact_name, phone, date]
                collect_billing_button.configure(bg="green")
                # checks if paramaters met to enable export button
                E = True
                if E and D and T and A:
                    print("All conditions true")
                    enable_export()

            # button to collect billing information and sore in a list
            collect_billing_button = Button(billing_frame, text="ENTER", bg="red", command=collect_billing_info)
            collect_billing_button.grid(row=3, column=0)

    # creates guts to work performed frame
    class WorkPerformedClass:
        def __init__(self):
            # create work performed label
            work_performed_label = Label(work_performed_frame, text="Work Performed")
            work_performed_label.grid(row=0, column=2)
            # creates location label
            location_label = Label(work_performed_frame, text="Location")
            location_label.grid(row=0, column=3)
            # creates hours label
            hours_label = Label(work_performed_frame, text="Hours")
            hours_label.grid(row=0, column=4)
            # work menu variable
            global clicked
            clicked = StringVar()
            clicked.set("Drive")
            # creates work menu drop down
            work_menu = OptionMenu(work_performed_frame, clicked, *work_keys)
            work_menu.grid(row=1, column=0, padx=15)

            # creates work performed fields
            global Hours
            global Work
            global Location
            Hours = {}
            Location = {}
            enter = {}
            Work = {}
            for x in range(10):
                enter["enter{0}".format(x)] = EnterWorkButtons(x)
                Work["Work{0}".format(x)] = WorkPerformedEntry(x)
                Location["Location{0}".format(x)] = LocationEntry(x)
                Hours["Hours{0}".format(x)] = HoursEntry(x)

    # button for work entry column
    class EnterWorkButtons:
        def __init__(self, index: int):
            self.enter_work_button = Button(work_performed_frame, text="Enter" + str(index),
                                            command=lambda: enter_work_from_dic(str(index)))
            self.enter_work_button.grid(row=1 + index, column=1)

    # Inserts work performed from work enter button
    def enter_work_from_dic(w):
        Work.get("Work{0}".format(w)).work_performed_entry.insert(END, work_dic.get(clicked.get()))

    # work performed column
    class WorkPerformedEntry:
        def __init__(self, index: int):
            self.work_performed_entry = Entry(work_performed_frame, text="work" + str(index), width=75)
            self.work_performed_entry.grid(row=1 + index, column=2)

    # collects data from all work performed entry fields and stores in list.
    def collect_work_performed():
        global work_performed_list
        work_performed_list = []
        for x in range(10):
            work_performed_list.append(Work["Work{0}".format(x)].work_performed_entry.get())

    # location entry column
    class LocationEntry:
        def __init__(self, index: int):
            self.location_entry = Entry(work_performed_frame, text="location" + str(index), width=13)
            self.location_entry.grid(row=1 + index, column=3)

    # collects all locations from entry fields
    def collect_locations():
        global location_list
        location_list = []
        for x in range(10):
            location_list.append(Location["Location{0}".format(x)].location_entry.get())

    # creates hours column
    class HoursEntry:
        def __init__(self, index: int):
            self.hours_entry = Entry(work_performed_frame, text="hours" + str(index), width=4)
            self.hours_entry.grid(row=1 + index, column=4)

    # totals up all the hours and prints to tallied hours label
    def tally_hours():
        global total_hours
        global hours_list
        global E
        global T
        global D
        global A
        total_hours = 0.0
        hours_list = []
        # if hours is left blank, replace with 0.
        for x in range(10):
            if Hours["Hours{0}".format(x)].hours_entry.get() == '':
                Hours["Hours{0}".format(x)].hours_entry.insert(0, 0)
            total_hours = total_hours + float(Hours["Hours{0}".format(x)].hours_entry.get())
            hours_list.append(Hours["Hours{0}".format(x)].hours_entry.get())
        collect_work_performed()
        collect_locations()
        # turns tall button green when pressed.
        tally_hours_button.configure(bg="green")

        # enables export button if conditions met
        T = True
        if E and D and T and A:
            enable_export()

        # display total hours when tally_hours button pressed.
        tallied_hours_label = Label(work_performed_frame, text="Hours" + str(total_hours))
        tallied_hours_label.grid(row=13, column=4)

    # button to tally up all the hours
    tally_hours_button = Button(work_performed_frame, text="Tally", bg="red", command=tally_hours)
    tally_hours_button.grid(row=12, column=4)

    # disables export button after use, and then exports ticket to excel.
    def export_ticket_button():
        export_button.config(state="disable")
        export_ticket()

    # ticket summary field
    class TicketSummary:
        def __init__(self):
            unit_label = Label(summary_frame, text="Unit")
            unit_label.grid(row=0, column=0)

            hours_label = Label(summary_frame, text="Hours")
            hours_label.grid(row=0, column=1)

            rate_label = Label(summary_frame, text="Rate")
            rate_label.grid(row=0, column=2)

            amount_label = Label(summary_frame, text="Amount")
            amount_label.grid(row=0, column=3)

            global Hours_S
            global Rate
            global Amount
            global Unit
            Hours_S = {}
            Amount = {}
            Rate = {}
            Unit = {}
            # builds summary fields
            for x in range(5):
                Unit["unit{0}".format(x)] = UnitEntry(x)
                Hours_S["hours{0}".format(x)] = HoursSummary(x)
                Rate["rate{0}".format(x)] = RateEntry(x)
                Amount["amount{0}".format(x)] = AmountEntry(x)
            Unit["unit0"].unit_entry.insert(0, "Steamer")

    class UnitEntry:
        def __init__(self, index=int):
            self.unit_entry = Entry(summary_frame, text="Unit" + str(index))
            self.unit_entry.grid(row=1 + index, column=0)

            # collects all data from summary fields.
            def summary_button_press():
                get_unit()
                get_rates()
                total_up_summary()
                hours_x_rate()
                summary_button.config(bg="green")
                # enables export button if conditions are met
                a = True
                if E and D and T and a:
                    enable_export()

            summary_button = Button(summary_frame, bg="red", text="Summary", command=summary_button_press)
            summary_button.grid(row=7, column=3)

    # hours summary field
    class HoursSummary:
        def __init__(self, index=int):
            self.hours_summary_entry = Entry(summary_frame, text="Hours" + str(index), width=4)
            self.hours_summary_entry.grid(row=1 + index, column=1)

    # stores hours from summary frame to a list
    def total_up_summary():
        global hours_summary
        hours_summary = []

        for x in range(5):
            hours_summary.append(Hours_S["hours{0}".format(x)].hours_summary_entry.get())
        print(hours_summary)

    # creates RateEntry fields
    class RateEntry:
        def __init__(self, index=int):
            self.rate_entry = Entry(summary_frame, text="Rate" + str(index), width=8)
            self.rate_entry.grid(row=1 + index, column=2)

    # stores units to list from summary frame
    def get_unit():
        global unit_list
        unit_list = []
        for x in range(5):
            unit_list.append(Unit["unit{0}".format(x)].unit_entry.get())

    # stores rates from summary frame to list.
    def get_rates():
        global rates_list
        rates_list = []
        for x in range(5):
            rates_list.append(Rate["rate{0}".format(x)].rate_entry.get())

        # multiplies hours by rate and inserts to field.
    def hours_x_rate():
        amount_list = []
        for x in range(5):
            Amount["amount{0}".format(x)].amount.delete(0, 5)
            if hours_summary[x] == '':
                hours_summary[x] = 0
            if rates_list[x] == '':
                rates_list[x] = 0

            if Amount["amount{0}".format(x)].amount.get() == '':
                Amount["amount{0}".format(x)].amount.insert(0, 0)
            amount_list.append(float(hours_summary[x]) * float(rates_list[x]))
            Amount["amount{0}".format(x)].amount.insert(0, amount_list[x])

    # creates amount entry fields
    class AmountEntry:
        def __init__(self, index=int):
            self.amount = Entry(summary_frame, text="Amount" + str(index), width=8)
            self.amount.grid(row=1 + index, column=3)

    class DriverField:
        def __init__(self):
            self.driver = Label(driver_frame, text="Driver:")
            self.driver.grid(row=4, column=0)

            def confirm_driver():
                global driver_s
                # Work.get("Work{0}".format(w)).work_performed_entry.insert(END, work_dic.get(clicked.get()))
                driver_s = driver_dic.get(driver_selected.get())
                driver_label = Label(driver_frame, text=driver_dic.get(driver_selected.get()))
                driver_label.grid(row=4, column=2)
                get_driver.config(bg="green")
                d = True
                if E and d and T and A:
                    enable_export()

            get_driver = Button(driver_frame, bg="red", text="Set Driver", command=confirm_driver)
            get_driver.grid(row=6, column=6)

            driver_selected = StringVar()
            driver_selected.set("Driver")

            driver_menu = OptionMenu(driver_frame, driver_selected, *driver_keys)
            driver_menu.grid(row=4, column=1)

    def enable_export():
        export_button.config(bg="green", state="normal")

    # exports all data to excel sheet Button
    export_button = Button(master_frame, bg="yellow", state="disable",
                           text="Export Ticket", command=export_ticket_button)
    export_button.grid(row=5, column=1)

    # collects data lists and writes to excel sheet.
    def export_ticket():
        ticket = pd.read_excel("NewTicket.xlsx")

        ticket.at["Charge To:", "DATA"] = charge_to
        ticket.at["Location:", "DATA"] = billing_information_list[1]  # work location
        ticket.at["Contact:", "DATA"] = billing_information_list[2]  # contact name
        ticket.at["Contact Phone:", "DATA"] = billing_information_list[3]  # phone number
        ticket.at["Date:", "DATA"] = str(billing_information_list[4])

        for x in range(10):
            ticket.at["Work{0}:".format(x), "DATA"] = work_performed_list[x]
            ticket.at["Location{0}:".format(x), "DATA"] = location_list[x]
            ticket.at["Hours{0}:".format(x), "DATA"] = hours_list[x]

        for x in range(5):
            ticket.at["Unit{}:".format(x), "DATA"] = unit_list[x]
            ticket.at["Hours_S{}:".format(x), "DATA"] = hours_summary[x]
            ticket.at["Rate{}:".format(x), "DATA"] = rates_list[x]
            # do amount calculation inside excel instead of this line
            # ticket.at["Amount{}:".format(x), "DATA"] = amount_list[x]
        ticket.at["Driver:", "DATA"] = driver_s
        ticket.at["Ticket Number", "DATA"] = ticket_counter

        # opens new ticket
        ticket_df = pd.DataFrame(ticket)
        wb = xw.Book("NewTicket.xlsx")

        ws = wb.sheets["Sheet"]
        # inserts data to excel sheet
        ws.range('A1').options(index=True).value = ticket_df
        # saves ticket to directory
        wb.save(sl + "\\SteamerTicket{}.xlsx".format(ticket_counter))
        print(ticket)

    # creates ticket frames
    bic = BillingInformationClass()
    wpc = WorkPerformedClass()
    ts = TicketSummary()
    D = DriverField()

    mainloop()


program()
