# easy-copy
Iterate over rows and copy columns from excel using keyboard keys. 

## Getting Started

This is just a little program that can be used to iterate over rows from an excel file and copy specified columns with the keyboard.
To run:
  ```
  sudo python3 EasyCopy.py --file <path to excel file> --sheet <sheet you want to work on> --columns <columns you want to copy>
  ```
Example/
  ```
  sudo python3 EasyCopy.py --file filename.xlsx --sheet 1 --columns ADE
  ```
The columns map to RIGHT-SHIFT, RIGHT-CONTROL and END. Iterating over rows is done with the left and right keys. 
So in the above example to copy column A you would press RIGHT-SHIFT and to copy column D you would press RIGHT-CONTROL.

pynput requires administrator priviliges to track keys hence the sudo command at the start. 


### Prerequisites

Python 3

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
