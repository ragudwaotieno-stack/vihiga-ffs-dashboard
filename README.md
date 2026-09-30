# FFS Tree Seedling Dashboard (Vihiga County / KEFRI)

Live dashboard that reads submissions straight from KoboToolbox.

- **KEFRI password:** totals only (county, sub-county, ward, FFS group by species). No farmer names or phones.
- **DoALF password:** everything above plus a Farmers tab and data checks.
- Excel download: County_by_species, FFS_groups, Wards, Sub_counties (+ Farmers for DoALF only).

## Deploy (about 15 minutes)
1. **Deploy the XLSForm in Kobo first** and upload `vihiga_logo.png` as its media file.
2. **Kobo token and form ID:** Account Settings > Security > API key. The form ID is the code in the form's web address (`/#/forms/<ID>/summary`).
3. **GitHub:** create a new **private** repository and upload every file in this folder (keep the `.streamlit` folder).
4. **Streamlit:** sign in at share.streamlit.io with GitHub > Create app > pick the repo, branch `main`, main file `app.py`.
5. Before deploying, open **Advanced settings > Secrets** and paste the values from `.streamlit/secrets.toml.example`, filled in.
6. Share the app link with KEFRI plus the KEFRI password. Keep the DoALF password to yourselves.

Without secrets the app runs in **demo mode** with sample data, so you can preview it first.

## Optional: reporting progress
Add `expected_groups.csv` to the repo with columns `ward,expected_groups` (ward names as in the form, e.g. `Shiru`). A progress table then appears under "Wards & sub-counties".

## Notes
- Duplicates (same ward + FFS group + farmer number) keep the latest submission. Untick the sidebar box to see everything.
- Group names are matched ignoring capital letters and extra spaces, so "Mbale FFS 1" and "MBALE FFS 1 " count as one group.
- Free Streamlit apps go to sleep after inactivity; opening the link wakes them in under a minute.
