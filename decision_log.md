# Decision Log

| No. | Decision | Alternatives considered | Reason | Status |
|-----|----------|-------------------------|--------|--------|
| 1 | Problem: district-level seasonal crop yield prediction | Price forecasting; pest/disease detection; farm-level yield | Free bulk data, useful output, achievable scope | Confirmed |
| 2 | Geographic scope: India, district level | Single state; global country-level | Best balance of detail and availability; may narrow to one state | Provisional |
| 3 | Start with one or two staple crops | All major crops | Keeps first pass manageable | Provisional |
| 4 | Target: yield (t/ha) plus at-risk flag | Production in tonnes only | Yield removes effect of changing area | Confirmed |
| 5 | Validation: temporal and spatial splits | Random train-test split | Random splits leak information | Confirmed |
| 6 | Public data only | Paid/proprietary data | Reproducibility and cost | Confirmed |
| 7 | Linear baselines, random forest, gradient boosting | Deep learning first | Strong on small tabular data; interpretable | Provisional |
