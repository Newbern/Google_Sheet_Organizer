from Google_Format import Sheet

# Setting up Spread sheet api and keys
Sheet = Sheet(key="https://docs.google.com/spreadsheets/d/--THIS RIGHT HERE--/", api='your_file.json')


# Money allowance for 2 weeks
def allowance(cell):
    # Getting numeric Value
    abc, cell_row, cell_col = Sheet.locate(cell)

    # Color list fill in
    colors = [["black", "yellow", "red", "red", "red", "green", "orange"],
              ["black"],
              ["black", "black", "black", "black"],
              ["black", "yellow", "yellow", "black"],
              ["black", "", "", "black"],
              ["black", "", "", "black"],
              ["black", "", "", "black"],
              ["black", "black", "black", "black", "black"]]
    Sheet.fill(cell, colors)

    # Set up gas sum
    abc, cell_row, cell_col = Sheet.locate(cell)
    gas_cell = f"{abc[cell_row + 2]}{cell_col + 1}"
    gas_total = f"{abc[cell_row + 4]}{cell_col + 4}:{abc[cell_row + 14]}{cell_col + 4}"
    Sheet.sum(gas_cell, gas_total)
    Sheet.money(gas_total)

    # Set up Food sum
    abc, cell_row, cell_col = Sheet.locate(cell)
    food_cell = f"{abc[cell_row + 3]}{cell_col + 1}"
    food_total = f"{abc[cell_row + 4]}{cell_col + 5}:{abc[cell_row + 14]}{cell_col + 5}"
    Sheet.sum(food_cell, food_total)
    Sheet.money(food_total)

    # Setup Extra sum
    abc, cell_row, cell_col = Sheet.locate(cell)
    extra_cell = f"{abc[cell_row + 4]}{cell_col + 1}"
    extra_total = f"{abc[cell_row + 4]}{cell_col + 6}:{abc[cell_row + 14]}{cell_col + 6}"
    Sheet.sum(extra_cell, extra_total)
    Sheet.money(extra_total)

    # Value list
    word = [
        ["", "Allowance", "Gas", "Food", "Extra", "Saved", "Added"],
        ["", 300],
        [],
        ["", "Allowance"],
        ["", "Gas", 180, "Gas"],
        ["", "Food", 80, "Food"],
        ["", "Extra", 25, "Extra"]
    ]

    # Making first row in list bold and centering
    Sheet.load(cell, word, 1, bold=True, center=True)

    # Adding Borders around first row
    Sheet.borders(f"{abc[cell_row + 1]}{cell_col}:{abc[cell_row + 5]}{cell_col}")

    # Making forth row in list bold and centering
    Sheet.load(cell, word, 4, bold=True, center=True)

    # Merging certain cells together
    Sheet.merge(f"{abc[cell_row + 1]}{cell_col + 3}:{abc[cell_row + 2]}{cell_col + 3}")


# Basic Money format for a year
def yearly_format(cell, name):
    # Getting location
    abc, cell_row, cell_col = Sheet.locate(cell)

    # Values list
    lst = [
        ["", name, "", "", "Extra", "Total month"],
        ["January", "", "", "", "", ""],
        ["February", "", "", "", ""],
        ["March", "", "", "", ""],
        ["April", "", "", "", ""],
        ["May", "", "", "", ""],
        ["June", "", "", "", ""],
        ["July", "", "", "", ""],
        ["August", "", "", "", ""],
        ["September", "", "", "", ""],
        ["October", "", "", "", ""],
        ["November", "", "", "", ""],
        ["December", "", "", "", ""]
    ]
    Sheet.load(cell, lst)

    # Making First row bold & center
    Sheet.load(cell, lst, 1, bold=True, center=True)

    # Background cell color
    color = [["", "", "", "", "", "green"]]
    Sheet.fill(cell, color)

    # Merging Cells together
    Sheet.merge(f"{abc[cell_row + 1]}{cell_col}:{abc[cell_row + 3]}{cell_col}")

    # Adding Borders to the cells
    Sheet.borders(f"{abc[cell_row + 1]}{cell_col}:{abc[cell_row + 4]}{cell_col}")


# Running both functions
def main(cell):
    # Getting the Cells location
    abc, cell_row, cell_col = Sheet.locate(cell)

    # Running first function
    yearly_format(cell, "Caleb")

    # Making sure this runs 13 spaces down from where the first function starts
    allowance(f"{abc[cell_row]}{cell_col + 13}")


# Running both functions
main("A1")
