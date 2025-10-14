# Yard Analysis System

A Python-based railway yard management analysis tool that processes outbound train schedules and yard plans to calculate car arrival patterns and dwell times.

## Overview

This system analyzes railway yard operations by:
- Processing departure schedules to identify trains and their block compositions
- Tracking car arrivals through a 24-hour period
- Calculating dwell times and car-hours for yard capacity planning
- Generating detailed reports for each outbound train

## Features

- **Automated Train Analysis**: Processes all trains from departure schedule automatically
- **Block Tracking**: Handles both dedicated block tracks and mixed SPARE tracks
- **Time-based Calculations**: 24-hour rolling analysis with midnight crossing support
- **Dwell Time Analysis**: Calculates car-hours for capacity and efficiency metrics
- **Excel Output**: Individual worksheets for each train with comprehensive metrics

## Requirements

```
pandas
openpyxl
```

Install dependencies:
```bash
pip install pandas openpyxl
```

## File Structure

```
project/
├── data/
│   ├── TH-Outbound-Train-Plan-2025.xlsx    # Departure schedule
│   └── alt_1.csv                            # Yard plan
├── results/
│   └── yard_analysis_results_alt_1.xlsx    # Output file
└── yard.py                                  # Main script
```

## Input Files

### 1. Departure Schedule (`TH-Outbound-Train-Plan-2025.xlsx`)

Required columns in `Worksheet1`:
- **Train**: Train identifier (e.g., ITHNAS, ITHEST)
- **Scheduled Departure**: Departure time
- **Bocks**: Comma-separated block types (e.g., "RLK, ESTR")

### 2. Yard Plan (`alt_1.csv`)

Required structure:
- **Time**: Time intervals (format: "0:00-0:15")
- **Pull [X]**: Pull track columns containing train identifiers
- **Block Columns**: Named columns for each block type
- **SPARE [1-5]**: Spare tracks with mixed blocks (format: "2 CHBR 1 CHG")

## How It Works

### Step 1: Extract Departure Information
- Reads all trains from departure schedule
- Identifies blocks assigned to each train
- Records scheduled departure times

### Step 2: Find Pull Operations
- Searches yard plan for earliest pull time of each train
- Handles midnight crossing scenarios (23:xx to 0:xx)
- Identifies the clearing operation row (excluded from arrival counts)

### Step 3: Calculate CAR_ARRIVING
For each hour (0-23):
- Counts cars in dedicated block columns
- Parses and counts cars in SPARE tracks containing target blocks
- Excludes the earliest pull time row (clearing operation)
- **Ensures non-negative values** (converts any negative counts to 0)

### Step 4: Compute Dwell Metrics

For each train, calculates:

1. **DWELL_HOURS**: Hours until next day midnight
   - 12 am → 24 hours
   - 23 pm → 1 hour

2. **CAR_ARRIVING × DWELL_HOURS**: Product of cars and dwell time

3. **TOTAL_CAR**: Sum of all CAR_ARRIVING

4. **TOTAL_CAR_HOURS**: Sum of all (CAR_ARRIVING × DWELL_HOURS)

5. **CAR_HOURS**: Cumulative car-hours (calculated backwards from 23:00)
   - Hour 23: `TOTAL_CAR_HOURS - TOTAL_CAR + 24 × CAR_ARRIVING(23)`
   - Hour 22-0: `Previous_CAR_HOURS - TOTAL_CAR + 24 × CAR_ARRIVING`

## Output Format

Excel file with one worksheet per train containing:

| Column | Description |
|--------|-------------|
| Train | Train identifier |
| Time | Hour (0:00 - 23:00) |
| CAR_ARRIVING | Number of cars arriving this hour |
| DWELL_HOURS | Hours until midnight (24 - hour) |
| CAR_ARRIVING_X_DWELL | CAR_ARRIVING × DWELL_HOURS |
| CAR_HOURS | Cumulative car-hours (backwards calculation) |
| TOTAL_CAR | Total cars for this train |
| TOTAL_CAR_HOURS | Total car-hours for this train |
| DEPARTURE_TIME | Scheduled departure time |

## Usage

```python
python yard_design
```

Or customize file paths:

```python
from yard_design import main

departure_file = "data/TH-Outbound-Train-Plan-2025.xlsx"
yard_plan_file = "data/alt_1.csv"
output_file = "results/yard_analysis_results_alt_1.xlsx"

main(departure_file, yard_plan_file, output_file)
```

## Example Output

For train ITHNAS (departing 03:30, blocks: EVL, NAS):

```
Processing train: ITHNAS
  Scheduled departure: 03:30
  Blocks: EVL, NAS
  Earliest pull time: 23:30-23:45 (hour: 23)
  Completed! Total Cars: 45, Total Car Hours: 789.00
```

## Key Features

### Midnight Crossing Handling
The system correctly identifies when train operations span across midnight (23:xx to 0:xx), ensuring the earliest pull time is accurately determined.

### SPARE Track Parsing
Automatically parses complex SPARE track formats:
- "2 CHBR 1 CHG" → 2 CHBR blocks + 1 CHG block
- "3 EVL" → 3 EVL blocks

### Negative Value Protection
All CAR_ARRIVING calculations are constrained to non-negative values. If a calculation results in a negative number, it is automatically converted to 0.

### Error Handling
- Gracefully handles missing data (NaN values)
- Skips trains not found in yard plan with warnings
- Creates output directories automatically
- Sanitizes sheet names for Excel compatibility

## Troubleshooting

### Train Not Found Warning
```
Warning: TRAIN_NAME not found in yard_plan
```
**Solution**: Verify train name matches exactly in both files (case-sensitive)

### No Valid Blocks Information
```
Warning: TRAIN_NAME has no valid blocks information
```
**Solution**: Check that the 'Bocks' column contains comma-separated block names

### Missing Columns
**Solution**: Ensure yard_plan.csv has:
- 'Time' column
- At least one 'Pull' prefixed column
- Block columns matching departure schedule

## Performance Notes

- Processing time depends on number of trains and yard plan size
- Typical processing: ~1-2 seconds per train
- Excel writing may take longer for large datasets

## Author Notes

This tool is designed for railway yard capacity planning and optimization. The backward calculation method for CAR_HOURS provides insights into cumulative yard occupancy over time, helping identify bottlenecks and optimize train scheduling.

---

**Version**: 1.0  
**Last Updated**: 2025-10-11