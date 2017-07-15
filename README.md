# game-of-life

[![MIT License](http://img.shields.io/badge/license-MIT-cccc00.svg?style=flat)](LICENSE.md)
[![Download](https://img.shields.io/badge/download-the&nbsp;xlsm&nbsp;file-00c0f0.svg?style=flat)](https://github.com/cygnan/game-of-life/raw/master/game-of-life.xlsm)

Implementation of Conway's Game of Life in Excel VBA.

![GIF](https://user-images.githubusercontent.com/25865313/27192246-942392f4-5235-11e7-9bb1-d1ad0f52fce4.gif)

## What's Conway's Game of Life?

It's a simulation game of real life processes. It follows the rules below.

### Rules

> All eight of the cells surrounding the current one are checked to see if they are on or not. Any cells that are on are counted, and this count is then used to determine what will happen to the current cell.
>
>1. Death: if the count is less than 2 or greater than 3, the current cell is switched off.
>1. Survival: if (a) the count is exactly 2, or (b) the count is exactly 3 and the current cell is on, the current cell is left unchanged.
>1. Birth: if the current cell is off and the count is exactly 3, the current cell is switched on.

Quoted from [Wolfram MathWorld](http://mathworld.wolfram.com/GameofLife.html).

## Usage

1. Open game-of-life.xlsm.
1. Then, push the "START" button to start a simulation. 

- Click or double click on cells to fill in light blue.
- Right click to empty. 
- You can select multiple cells. If the cell at the upper-left of the selected range was off, all of the selected cells will be switched on. Otherwise, switched off.
- Push the "CLEAR" button to empty all of the cells.
- The rules can be changed by switching on or off the checkboxes.
- If the "PROTECTED VIEW" alert is displayed, push the "Enable Editing" button.
- If the "SECURITY WARNING" alert is also displayed, push the "Enable Content" button.

## Contribution

1. Fork it
1. Create your feature branch
1. Commit your changes
1. Push to the branch
1. Create a new Pull Request

## License

Copyright &copy; 2017 Cygnan  
Licensed under the MIT License, see [LICENSE.md](LICENSE.md).
