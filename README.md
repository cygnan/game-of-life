# game-of-life
An implementation of Conway's Game of Life in Excel VBA.

## Usage
1. Open game-of-life.xlsm.
1. If "SECURITY WARNING" is displayed, push the "Enable Content" button.
1. Then, push the "START" button to start a simulation. 

- Click or double click on cells to fill in light blue.
- Right click to empty. 
- You can select multiple cells. If the cell at the upper-left of the selected range is off, all of the selected cells is switched on. Otherwise, switched off.
- Push the "CLEAR" button to empty all of the cells.
- The rules can be changed by switching on or off the checkboxes.

## What's Conway's Game of Life?
It's a simulation game of real life processes. It follows the rules below.

### Rules
> All eight of the cells surrounding the current one are checked to see if they are on or not. Any cells that are on are counted, and this count is then used to determine what will happen to the current cell.
>
>1. Death: if the count is less than 2 or greater than 3, the current cell is switched off.
>1. Survival: if (a) the count is exactly 2, or (b) the count is exactly 3 and the current cell is on, the current cell is left unchanged.
>1. Birth: if the current cell is off and the count is exactly 3, the current cell is switched on.

Quoted from [Wolfram MathWorld](http://mathworld.wolfram.com/GameofLife.html).

## License
Copyright &copy; 2017 Cygnan  
Licensed under the MIT License, see [LICENSE.md](LICENSE.md).
