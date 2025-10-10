import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def time_to_angle(time_value):
    """Convert time (datetime or string) into angle (radians) for 24-hour polar chart."""
    if pd.isnull(time_value):
        return np.nan
    if isinstance(time_value, str):
        try:
            time_value = pd.to_datetime(time_value).time()
        except:
            return np.nan
    return 2 * np.pi * (time_value.hour + time_value.minute / 60) / 24


def plot_train_chart(inbound_file, outbound_file, save_path="yard_train_chart.png"):
    # === Read inbound data ===
    inbound_df = pd.read_excel(inbound_file, sheet_name="Schedule", usecols=["Train", "Scheduled Arrival"])
    inbound_df = inbound_df.dropna(subset=["Train", "Scheduled Arrival"])
    inbound_df["Angle"] = inbound_df["Scheduled Arrival"].apply(time_to_angle)
    inbound_df["Type"] = "Inbound"

    # === Read outbound data ===
    outbound_df = pd.read_csv(outbound_file, usecols=["Departure Train", "Departure Time"])
    outbound_df = outbound_df.rename(columns={"Departure Train": "Train", "Departure Time": "Time"})
    outbound_df = outbound_df.dropna(subset=["Train", "Time"])
    outbound_df["Angle"] = outbound_df["Time"].apply(time_to_angle)
    outbound_df["Type"] = "Outbound"

    # === Polar figure setup ===
    fig, ax = plt.subplots(figsize=(9, 9), subplot_kw={'projection': 'polar'})
    ax.set_theta_direction(-1)
    ax.set_theta_zero_location('N')
    ax.set_yticklabels([])
    ax.set_xticks(2 * np.pi * np.arange(0, 24, 1) / 24)
    ax.set_xticklabels(np.arange(0, 24, 1), fontsize=10)

    # === Yard circle (smaller and gray) ===
    yard_radius = 0.15
    yard_circle = plt.Circle((0, 0), yard_radius, color='lightgray', alpha=0.9)
    ax.add_artist(yard_circle)
    ax.text(0, 0, "Yard", color="black", ha="center", va="center", fontsize=10, fontweight='bold')

    # === Plot inbound trains (red arrows pointing inward) ===
    for _, row in inbound_df.iterrows():
        angle = row["Angle"]
        if np.isnan(angle):
            continue
        ax.annotate("",
                    xy=(angle, yard_radius),     # head at center edge
                    xytext=(angle, 1.0),         # tail at outer ring
                    arrowprops=dict(arrowstyle="->", color="red", lw=1.5))
        # Slight radial offset to reduce label overlap
        label_radius = 1.05 + 0.03 * np.random.randn()
        ax.text(angle, label_radius, row["Train"], color="red", fontsize=8,
                ha="center", va="center", rotation=np.degrees(-angle),
                rotation_mode='anchor')

    # === Plot outbound trains (green arrows pointing outward) ===
    for _, row in outbound_df.iterrows():
        angle = row["Angle"]
        if np.isnan(angle):
            continue
        ax.annotate("",
                    xy=(angle, 1.0),             # head at outer ring
                    xytext=(angle, yard_radius), # tail near center
                    arrowprops=dict(arrowstyle="->", color="green", lw=1.5))
        # Slight radial offset to reduce label overlap
        label_radius = 1.1 + 0.03 * np.random.randn()
        ax.text(angle, label_radius, row["Train"], color="green", fontsize=8,
                ha="center", va="center", rotation=np.degrees(-angle),
                rotation_mode='anchor')

    # === Style settings ===
    ax.set_title("24-hour Yard Chart",
                 fontsize=14, pad=25, fontweight='bold')
    ax.grid(alpha=0.3)

    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.show()
    print(f"Chart saved to {save_path}")


# ==== Run example ====
if __name__ == "__main__":
    plot_train_chart(
        inbound_file="data/TH-Inbound-Train-Plan-2025.xlsx",
        outbound_file="data/alt_1_depart.csv",
        save_path="results/yard_train_chart_refined.png"
    )
