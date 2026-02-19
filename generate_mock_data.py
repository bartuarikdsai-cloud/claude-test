import json
import numpy as np
import pandas as pd

np.random.seed(42)

N = 10_000

# --- Gender ---
gender = np.random.choice(["Male", "Female"], size=N, p=[0.52, 0.48])

# --- Age (18-75, skewed toward 30-50) ---
age = np.clip(np.random.normal(loc=40, scale=12, size=N), 18, 75).astype(int)

# --- Car Model Year (2000-2025, weighted toward newer) ---
years = np.arange(2000, 2026)
weights = np.linspace(1, 5, len(years))  # newer cars more likely
weights /= weights.sum()
car_model_year = np.random.choice(years, size=N, p=weights)

car_age = 2025 - car_model_year

# --- Annual Premium ---
# Base premium ~$1,200, adjusted by age, gender, car age
base_premium = 1200

# Younger drivers pay more (U-shaped: young and very old pay more)
age_factor = np.where(age < 25, 1.45,
             np.where(age < 30, 1.15,
             np.where(age < 60, 1.0,
             np.where(age < 70, 1.10, 1.25))))

# Males pay slightly more
gender_factor = np.where(gender == "Male", 1.08, 1.0)

# Older cars: slightly higher premium
car_age_factor = 1.0 + car_age * 0.012

# Add noise
noise = np.random.normal(1.0, 0.10, size=N)

annual_premium = base_premium * age_factor * gender_factor * car_age_factor * noise
annual_premium = np.clip(annual_premium, 500, 5000).round(2)

# --- Total Loss (claims) ---
# ~70% of customers have zero claims
# Claim probability higher for young drivers and old cars
claim_base_prob = 0.28
age_claim_adj = np.where(age < 25, 0.15,
                np.where(age < 30, 0.05,
                np.where(age < 60, 0.0,
                np.where(age < 70, 0.05, 0.10))))
car_claim_adj = car_age * 0.005

claim_prob = np.clip(claim_base_prob + age_claim_adj + car_claim_adj, 0.05, 0.70)
has_claim = np.random.binomial(1, claim_prob)

# Claim amounts: lognormal distribution (most claims small, some very large)
claim_amount = np.where(
    has_claim == 1,
    np.random.lognormal(mean=7.5, sigma=1.0, size=N),  # median ~$1,800
    0.0
)
total_loss = np.clip(claim_amount, 0, 80_000).round(2)

# --- Loss Ratio ---
loss_ratio = np.where(annual_premium > 0, total_loss / annual_premium, 0.0).round(4)

# --- Build DataFrame ---
df = pd.DataFrame({
    "customer_id": range(1, N + 1),
    "gender": gender,
    "age": age,
    "car_model_year": car_model_year,
    "annual_premium": annual_premium,
    "total_loss": total_loss,
    "loss_ratio": loss_ratio,
})

# --- Save CSV/Excel ---
df.to_csv("auto_insurance_data.csv", index=False)
df.to_excel("auto_insurance_data.xlsx", index=False)

# --- Save compact JSON for dashboard embedding ---
# Use short keys to minimize file size: i=id, g=gender, a=age, y=car_year, p=premium, l=loss
records = []
for _, row in df.iterrows():
    records.append([
        int(row["customer_id"]),
        1 if row["gender"] == "Male" else 0,  # 1=Male, 0=Female
        int(row["age"]),
        int(row["car_model_year"]),
        round(row["annual_premium"], 2),
        round(row["total_loss"], 2),
    ])

with open("dashboard_data.json", "w") as f:
    json.dump(records, f, separators=(",", ":"))

print(f"Generated {len(df)} rows")
print(f"dashboard_data.json: {len(json.dumps(records, separators=(',', ':')))/1024:.0f} KB")
print(f"\n--- Summary ---")
print(f"Claim rate: {(total_loss > 0).mean():.1%}")
print(f"Avg premium: ${annual_premium.mean():,.0f}")
print(f"Avg loss (all): ${total_loss.mean():,.0f}")
print(f"Avg loss (claimants only): ${total_loss[total_loss > 0].mean():,.0f}")
