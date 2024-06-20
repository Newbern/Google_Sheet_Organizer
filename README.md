# Google Sheets Basic Format
## Modules
* gspread
* oauth2client

This is just a shorter and easier way to create Google spreadsheet files for anyone doing any Google Sheet function of the sort.

I made all this easy to understand and shortened the work load on a user, I made all these doable using a class where
all you do is just enter the api credentials and the url key to give the code access to.

# Setting up the api and gaining access
```bash
class Sheet:
    # The api is the credentials you downloaded from Google
    # and The key is the Google url key used to get access to the spreadsheet
    def __init__(self, api, key):
        # Define the scope
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]

        # Getting information from the api file
        credentials = ServiceAccountCredentials.from_json_keyfile_name(api, scope)

        # Authorize the client for Accesses into the Sheet
        client = gspread.authorize(credentials)

        # Having the correct Google sheet url to edit in
        spreadsheet = client.open_by_key(key)

        # This is the official Spread Sheet
        self.sheet = spreadsheet

        # First Sheet being edit inside the SpreadSheet
        self.worksheet = spreadsheet.sheet1
```
This is giving us access to the spreadsheet, so we can do our own creative projects

# Writing in Cell & Filling in the background
to make this easy First Color
```bash
# Getting the color by string or tuple
    @staticmethod
    def get_color(rbg: str or tuple) -> tuple:
        if rbg is str:
            rbg.lower()
        # Green
        if rbg == "green":
            rbg = (0, 255, 0)
        # Red
        elif rbg == "red":
            rbg = (255, 0, 0)
        # Yellow
        elif rbg == "yellow":
            rbg = (255, 255, 0)
        # Orange
        elif rbg == "orange":
            rbg = (255, 153, 0)
        # Pink
        elif rbg == "pink":
            rbg = (255, 0, 255)
        # Blue
        elif rbg == "blue":
            rbg = (0, 0, 255)
        # Black
        elif rbg == "black":
            rbg = (0, 0, 0)
        # White
        elif rbg == "white" or rbg == "":
            rbg = (255, 255, 255)
        # No color selected
        else:
            rbg = (0, 0, 0)

        # just because I like the simple (255,255,255) you are expose to use (1, 0.2, 0)
        rbg = (rbg[0] / 255, rbg[1] / 255, rbg[2] / 255)
        return rbg
```
What this does is whenever you write the color it will shoot out that tuple which the cell can use to change the color 
to text for background it won't matter.

```bash
# Cell Background
    def fill(self, cell, color) -> None:
        # Getting the color
        # Checking to see if this is a list and if it is
        # You can fill in multiple colors as once
        if isinstance(color, list):
            formats_lst = []
            abc, cell_x, cell_y = self.locate(cell)
            y = 0
            for col in color:
                x = 0
                for color in col:
                    cell = f"{abc[cell_x + x]}{cell_y + y}"
                    
                    # Getting the color
                    color = self.get_color(color)
                    
                    # Setting up method to be used in batch_format
                    update = {
                        "range": cell,
                        "format": {
                            "backgroundColor": {
                                "red": color[0],
                                "green": color[1],
                                "blue": color[2]
                            }
                        }
                    }
                    
                    formats_lst.append(update)
                    x += 1
                y += 1
            # Once for loop finishes getting all the colors it will update it all the cells at once
            self.worksheet.batch_format(formats_lst)

        # if this is not a list it will run as a str, tuple and so on
        else:
            # Getting the color
            color = self.get_color(color)

            # Filling in the cell background
            self.worksheet.format(cell, {
                "backgroundColor": {
                    "red": color[0],
                    "green": color[1],
                    "blue": color[2]
                }})
```
This will fill in the background of a cell regardless if its one cell or one-thousand all different colors.

Now sense we are at big uploads like this filling in colors lets talk about upload small and large text values.
```bash
# Cell Writing and TextColor changing and Bold
    def write(self, cell, text=False, bold=False, color=False) -> None:

        # Writing text for the cell
        self.worksheet.update_acell(cell, text)

        # Getting color
        color = self.get_color(color)

        self.worksheet.format(cell, {
            "textFormat": {
                "bold": bold,
                "foregroundColor": {
                    "red": color[0],
                    "green": color[1],
                    "blue": color[2]
                }
            }
        })
```
im sure that was easy enough as it was written, but now you can easily bold and change the text color

# Uploading Large values
```bash
   # Cell Uploading List
    def load(self, cell, lst, lst_spc: int = False, bold=False, color=False, center=False) -> None:
        # Good for loading list of values
        self.worksheet.update(cell, lst)

        # Getting color
        color = self.get_color(color)

        # Getting specific cell location
        abc, cell_x, cell_y = self.locate(cell)

        # Getting the complete list cell beginning and ending
        if lst_spc:
            lst_spc -= 1
            cell = f"{abc[cell_x]}{cell_y + lst_spc}:{abc[cell_x + len(lst[lst_spc]) - 1]}{cell_y + lst_spc}"
        else:
            # Getting full list length
            cell = f"{cell}:{abc[cell_x + len(max(lst, key=len))]}{cell_y + len(lst)}"

        # Changing color and boldness
        self.worksheet.format(cell, {
            "textFormat": {
                "bold": bold,
                "foregroundColor": {
                    "red": color[0],
                    "green": color[1],
                    "blue": color[2]
                }
            }
        })

        if center:
            if isinstance(center, bool):
                center = "center"
            self.center(cell, center)
```
Now this can load a large set of values in one push, and it can change a specific rows boldness or text color
it can also center the text if that is what one is thinking.

# Sum of multiple cells
```bash
# Sum
    def sum(self, cell, value: str or tuple) -> None:
        # if there is a bunch of cells you want to add together this will add them
        if isinstance(value, tuple):
            value = ",".join(value)
            
        self.write(cell, f"=sum({value})")
```
This line of code `=sum(A1:A5)` will sum up the numbers of those cells I wrote this line of code just to help make things run a bit easier.

# Locating cells, Finding there Value & Getting the cells Value
## Locate a cell
So first ill share with you what I mean by "Locating cells".
```bash
# Cell Number
    @staticmethod
    def locate(cell: str) -> tuple:
        # Cells row
        cell_row = cell[0].upper()

        # Cells column
        cell_col = cell[1:]

        # Abc to find the location number of Cell row
        abc = "".join([chr(64 + i) for i in range(1, 27)])
        return abc, abc.find(cell_row), int(cell_col)
```
As you can see this staticmethod returns a few things.

First `abc`, All this does is return a list filled with the alphabet I did this simple just to have it on hand and to
save myself of constantly using `chr`. I want it to be easy as 1,2,3. plus it just simplifies things.

Next what `abc.find(cell_row)`, All it does is go through that list and find the Number it is and returns it back 
as a `int`
the reason I did that is, so you can just add and find yourself at a different row when you do `abc[cell_row + 1]` now 
it would be `B`.

Lastly, you just serious bring back the cell number at the end and all I did want make it an integer just in case you want 
to change that one as well.

Then we will talk about Finding a cell just by its Value.

## Finding a cell by its Value
```bash
# Finding cell location
    def find(self, text: str) -> str:
        # Getting the value of all the cells
        all_cells = self.worksheet.get_all_values()
        
        # Using the enumerate function to keep up with the index and row
        for row_index, row in enumerate(all_cells, start=1):
            # Using it again to get the Index of the col to get the right Letter
            for col_index, col in enumerate(row, start=1):
                
                # Checking to see if This is the correct value
                if col == text:
                    # returning the row and col with chr
                    # chr is just a built-in module with letters and numbers
                    # you don't hit the alphabet until 64 characters in
                    return f"{chr(64 + col_index)}{row_index}"
```
Just to put this in basic words all this does is find the exact cell with the value in it, and it will return it just
like it would if you where to use the locate function

## Getting the Value from a specific cell
```bash
# Getting Value
    def get(self, cell: str) -> None:
        # Returning value of the cell
        return self.worksheet.get(cell)
```
Easy-peasy all this does is return the value easiest thing ever

# Clear
```bash
# Clear
    def clear(self, cell: str = False) -> None:
        # Checks to see if you want a specific cell erased or all of it
        if cell:
            self.worksheet.batch_clear([cell])
        else:
            # whole sheet
            cell = "1:1000"
            self.worksheet.clear()

        # Erasing the Text color and Bolding
        # Erasing the Background color
        # Resetting horizontalAlignment, Boarders & NumberFormat
        self.worksheet.format(cell, {
            "textFormat": {
                "bold": False,
                "foregroundColor": {
                    "red": 0,
                    "green": 0,
                    "blue": 0
                }
            },
            "backgroundColor": {
                "red": 1,
                "green": 1,
                "blue": 1
            },
            "horizontalAlignment": None,
            "borders": None,
            "numberFormat": None
        })

        # Unmerging All Cells
        self.worksheet.unmerge_cells(cell)
```
This Erases everything that I have set functions to and of course all of these can be changed to the user's preference.

The rest of are basically just like the modules suggests to use it all I did was shorten the steps to run them.
