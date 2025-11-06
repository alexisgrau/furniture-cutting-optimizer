---

# Furniture Cutting and Purchase Optimizer

This Node.js application reads a list of furniture parts from an Excel file, optimizes the cutting layout across available boards, and generates a detailed HTML report with cut plans, efficiency statistics, and a shopping list for board purchasing.

---

## Features

- Reads input data from an Excel file (`input.xlsx`)  
- Automatically groups parts by board thickness  
- Optimizes the cutting layout for each thickness  
- Generates **SVG cutting plans** with part names and dimensions  
- Produces an **HTML report** summarizing:
  - Number of boards required  
  - Total cost  
  - Material efficiency  
  - Shopping list (by board type)  
- Allows manual confirmation before processing  
- Simple configuration via the `CONFIG` object

---

## Requirements

- [Node.js](https://nodejs.org/) (v16 or higher)
- npm packages:
  - `xlsx`
  - `fs`
  - `path`
  - `readline` (native in Node.js)

Install dependencies with:

```bash
npm install
```

---

## File Structure

```
.
├── index.js        # Main program file
├── input.xlsx      # Excel file containing part dimensions
└── output/
    └── results.html  # Generated HTML report (created after execution)
```

---

## Input File Format (`input.xlsx`)

The program expects an Excel file (`input.xlsx`) with the following columns:

| Name | Dimensions | Thickness | Quantity | m² |
|--------:|------------------------|-------------|
|suppport haut | 470 mm	x	1200 mm |	16 mm |	1 |	0.564 |
|bord gauche | 470 mm	x	900 mm |16 mm |	1 |	0.423 |
|bord droit | 470 mm	x	900 mm |16 mm |	1 |	0.423 |

The script automatically skips empty rows and handles multiple quantities by duplicating entries.
I use the woodcutting plugin for FreeCAD to generate this file but you can create your own, based on same model.

---

## Configuration

The configuration is defined inside `index.js` in the `CONFIG` object:

```js
const CONFIG = {
  inputFile: 'input.xlsx',
  outputDir: 'output',
  boards: [
    {
      width: 2000,
      height: 500,
      thickness: 16,
      price: 10.9
    }
  ],
  kerf: 3,     // Cut width (mm)
  margin: 4    // Safety margin (mm)
};
```

You can define multiple board types by adding more entries to the `boards` array (for different thicknesses, sizes, or prices).

---

## Usage

1. Place your `input.xlsx` file in the same folder as `index.js`.
2. Run the program:

   ```bash
   node index.js
   ```

3. Review the data preview shown in the console.
4. Confirm by typing `Y` when prompted.
5. After processing, the tool will:
   - Display summary statistics in the terminal
   - Generate an HTML report located at `output/results.html`

Open `results.html` in your browser to visualize the optimized cutting plans.

---

## HTML Report Overview

The generated `results.html` file includes:

- **Shopping List**  
  Total boards required per type and cost breakdown

- **Summary Section**  
  Displays number of boards, total cost, total pieces, and material efficiency percentage

- **Detailed Cutting Plans**  
  For each board:
  - SVG visualization of piece placement  
  - List of cut dimensions and rotation indicators  
  - Individual usage rate (%)

You can print the report directly from your browser (a print button is included).

---

## How It Works

1. **Excel Parsing**  
   Reads part names and dimensions from the first sheet.

2. **Filtering by Thickness**  
   Excludes pieces that don’t match any available board thickness in `CONFIG.boards`.

3. **Placement Optimization**  
   Uses a simple heuristic packing algorithm to fit pieces on available boards with kerf and margin considerations.

4. **Visualization**  
   Generates scalable SVG drawings for each board layout.

5. **Report Generation**  
   Combines all data into a single, self-contained HTML report.

---

## Example Output

- **Boards required:** 3  
- **Total cost:** €32.70  
- **Efficiency:** 84.5%  
- **Pieces:** 45  

The `output/results.html` file displays interactive cut layouts for each board.

---

## Notes

- Only the first sheet of the Excel file is processed.  
- The program currently uses a **basic greedy placement algorithm**, not a full optimization solver.  
- Efficiency can be improved by refining `kerf`, `margin`, and board dimensions.  
- Large datasets may take longer to compute due to layout iteration.

---

## License

This project is open-source and provided for educational and personal use.  
You are free to modify and adapt it for your own furniture or carpentry projects.

---
