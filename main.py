import pandas as pd
import os


def process_train_data(departure_file, yard_plan_file, output_dir):
    dep_df = pd.read_excel(departure_file, sheet_name='Worksheet1')
    dep_df = dep_df[['Train', 'Scheduled Departure', 'Blocks']].dropna()

    yard_sheet1 = pd.read_excel(yard_plan_file, sheet_name='Sheet1')
    yard_sheet2 = pd.read_excel(yard_plan_file, sheet_name='Sheet2')

    os.makedirs(output_dir, exist_ok=True)

    for _, row in dep_df.iterrows():
        train_name = str(row['Train']).strip()
        dep_time = row['Scheduled Departure']
        blocks = [b.strip() for b in str(row['Blocks']).split(',')]

        print(f"\nProcessing train: {train_name} | Departure: {dep_time} | Blocks: {blocks}")

        # find the dept Train from Sheet2 of yard_plan
        matched_rows = yard_sheet2[yard_sheet2.astype(str).apply(lambda x: x.str.contains(train_name)).any(axis=1)]
        if matched_rows.empty:
            print(f"No match found for train {train_name} in yard_plan Sheet2.")
            continue

        # find block (assume first row is block & names）
        block_row = yard_sheet2.iloc[0]
        block_columns = [col for col in yard_sheet2.columns if any(b in str(block_row[col]) for b in blocks)]

        if not block_columns:
            print(f"No matching blocks {blocks} found in yard_plan for {train_name}.")
            continue

        # calculate CAR ARRIVING
        hours = list(range(0, 24))
        car_arriving_per_hour = []
        for h in hours:
            # todo: yard_plan structure check
            hour_rows = yard_sheet1[yard_sheet1['Hour'] == h] if 'Hour' in yard_sheet1.columns else yard_sheet1
            car_sum = hour_rows[block_columns].sum().sum()
            car_arriving_per_hour.append(car_sum)

        hourly_df = pd.DataFrame({
            'Train': [train_name] * 24,
            'Hour': hours,
            'CAR_ARRIVING': car_arriving_per_hour
        })

        # agg -> total_car, dwell_hours, total_car_hours, car_hours ===
        hourly_df['DWELL_HOURS'] = 24 - hourly_df['Hour']
        hourly_df['CARxDWELL'] = hourly_df['CAR_ARRIVING'] * hourly_df['DWELL_HOURS']

        total_car = hourly_df['CAR_ARRIVING'].sum()
        total_car_hours = hourly_df['CARxDWELL'].sum()

        # optimization table
        car_hours = []
        for i in reversed(range(24)):
            hour = hours[i]
            car_i = hourly_df.loc[hourly_df['Hour'] == hour, 'CAR_ARRIVING'].values[0]
            if i == 23:
                ch = total_car_hours - total_car + 24 * car_i
            else:
                ch = car_hours[-1] - total_car_hours - 24 * car_i
            car_hours.append(ch)
        car_hours = list(reversed(car_hours))
        hourly_df['CAR_HOURS'] = car_hours

        output_path = os.path.join(output_dir, f"{train_name}_summary.csv")
        hourly_df.to_csv(output_path, index=False)
        print(f"Saved {output_path}")

    print("\nAll trains processed successfully.")


if __name__ == "__main__":
    process_train_data(
        departure_file="data/TH-Outbound-Train-Plan-2025.xlsx",
        yard_plan_file="data/alt_1.xlsx",
        output_dir="results/train_summary_alt_1"
    )
