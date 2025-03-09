# GDNC Invoice Automation System

## Description

Tkinter-based app for generating invoice for "Giron du Nord 2025 à Concise" sponsoring based on selection and templates.

## Resources

- YouTube video "[NeuralNine - Invoice Automation System in Python - Full Project](https://youtu.be/JuBEC1RW8nA?si=-a1BploFfwDJsV0a)".

## Getting started

### Dependencies

- Tested on macOS Ventura version 13.4
- Python 3.10

### Environment

- Miniconda virtual environment: `automation-env` (Python 3.10).

## Executing program

TODO: Maybe update below!!!

```bash
conda activate automation-env
python main.py
```

## Development timeline (to open with [Markwhen](https://markwhen.com/) extension)

#fundamentals: #dbde33
#ui: #36d399
#backend: #f013ec

section GDNC Invoice Automation System #yellow

2024-09-04: Project launched! 🚀🏁

### Fundamentals

group Fundamentals #fundamentals

2024-09-04/2024-10-17: Follow tutorial from YouTube video "[NeuralNine - Invoice Automation System in Python - Full Project](https://youtu.be/JuBEC1RW8nA?si=-a1BploFfwDJsV0a)".
2024-10-17: Find a way to have pre-filled values (for dates and invoice number for instance) that the user can eventually manually change or leave as it is if correct (cf.: "GeeksforGeeks - How to Set the Default Text of Tkinter Entry Widget?" (https://www.geeksforgeeks.org/how-to-set-the-default-text-of-tkinter-entry-widget/)).
2024-10-17: Add check at the beginning of `create_invoice` to make sure that none of the fields are left empty.
2025-01-24: For the product to select, add a numeric field for being able to choose the quantity! ~~This should not be part of the dictionary maybe...~~
TODO: Create an Alfred workflow "GI" (GDNC Invoice) for launching the app.

endGroup

### UI

group UI #ui

2025-02-20: Add a "x" button next to each additional entry to eventually remove it if the user wants to.
2025-02-20: Add a scrolling bar to make sure I can scroll down to see all the products if there are too many (cf.: "Stack Overflow - Create a scrollbar to a full window tkinter in python" (https://stackoverflow.com/a/71682458)).
2025-02-20: Make sure the price fields are NOT editable by the user!
2025-02-21: Insert release version number at the top.
2025-02-21: Add checkboxes to make sure to be able select only default or only custom products or both.
2025-02-22/2025-02-26: Implement custom product rows (listen to entry changes in tkinter, cf. "[Stack Overflow - How do I get an event callback when a Tkinter Entry widget is modified?](https://stackoverflow.com/a/44365434)").
2025-02-26: Make "Spinbox" elements non editable by hand (i.e., set them as "readonly")
2025-02-26: Make sure toggle on and off default and custom product list coincides with `selected_default_product_dict` and `selected_custom_product_dict`
2025-02-26/2025-02-27: ~~Place GDNC PNG logo at the top~~ Implement spinning wheel with GDNC GIF logo.
2025-02-27: Make "title" field as "Spinbox" to select only either "Monsieur" or "Madame".
2025-03-05: ~~Add notification after pushing "generate invoice" button to tell the user to check outputs in the terminal~~ Add a label box next to GDNC logo to tell to the user to check outputs in the terminal and indicate at which stage of the generation we currently are.
2025-03-05: Once invoice has been generated, add static image similar to GDNC logo but with check mark emoji on top to indicate that process has finished.
TODO: Add "comment" entry (to be able to add comments related to the sponsor) at the bottom of the UI, right after the custom product frame.
TODO: Add checkboxes right before default product frame for sponsoring type: cash or material → Then also have those 2 columns in the database: Cash, Material with a symbol `True` to indicate selection.

endGroup

### Backend

group Backend #backend

2025-02-20: Add keyboard interrupt: So that when we hit Ctrl+C from the terminal, it quits the program UI and exits. Cmd + W might work and be sufficient! (This seems to work from itself already!)
2025-02-21: Automatic computation of total cost (by fetching selected product details, i.e. names, quantities and hence prices).
2025-02-21: Make more complex the script for being able to select more than 1 product → Maybe have several DOCX templates with from 1 to 5 products → Add a numerical field for selecting the number of products to add to the invoice, i.e., the number of product dropdown list to then DYNAMICALLY add to the UI (if possible with Tkinter!).
2025-02-27: Make sure overall product number (default + custom products) is ≤ 5 (depending on max products available on existing DOCX templates).
2025-03-03: Fetching selected sponsor information and create SponsorObject class.
2025-03-03: Add check for internet connection (since required for converting DOCX to PDF)
2025-03-05: Fill respected fields in the adapted DOCX template (depending on the number of products).
2025-03-05: Fix the font of the invoice number, and date (maybe in the DOCX document directly)
2025-03-05: At the very end, improve the script to automatically populate the corresponding Excel file listing sponsor data.
2025-03-05: Update and generate sponsor data Excel file with data of new sponsor entered in the app here.
2025-03-05: Products have to be stored in reference Excel file under the form of ~~list~~ dict ~~(both default and custom in the same list: default 1, custom 1, custom 2, etc.) → Also, add columns "num default products" and "num custom products"~~.
TODO: Implement nice logging message instead of simple print statements.

TODO: In case we have only custom products when generating the invoice, which corresponds to a "Donation" make sure to have a template extra for donations where ~~we do NOT have the field "enterprise name" +~~ we do NOT include TVA → Remove the TVA part!
TODO: Refactor the main parts of `create_invoice()` by grouping code in functions.
TODO: Introduce constant SKIP_INVOICE_GENERATION to skip docx generation and pdf conversion parts.
TODO: Display elapsed time for generating invoice in success pop-up appearing at the end of process.
TODO: Untrig GDNC checked logo and clear status label when clicking OK success pop-up appearing at the end of process.
TODO: Properly handle Invoice number → Default value from database, but can be intentionally left empty in UI (e.g., the invoice for Duckert where we won’t send the invoice to the company since they only gave material sponsoring and won’t give us some money but we generate a fake invoice and an entry in the database in order to keep a trace) → then anyway if not a number → Set `None` in database.
TODO: Idea for future app, maybe based on Streamlit, for analysing sponsoring data: Bar chart with chronological income from sponsoring, pie chart with quantity of the different sponsoring packs sold, distribution chart with 10 bins representing the radius distance from the sponsor location to Concise in [km] and a map with location point where we got sponsors from (e.g., "[Plotly Python Open Source Graphing Library Maps](https://plotly.com/python/maps/)") → Even better would be a 3D geographic map with 3D bar chart indicating the number of sponsors per location and also the amount of money obtained at each location → Even better would be to have a slider to see how this evolved across time! → Resources for 3D maps in Python:
    - https://docs.streamlit.io/develop/api-reference/charts/st.pydeck_chart
    - https://community.plotly.com/t/is-it-possible-to-make-this-with-plotly/73696
    - https://medium.com/plotly/5-awesome-tools-to-power-your-geospatial-dash-app-c71ae536750d
    - https://youtu.be/Bne9VASvDoI?si=DW5W0QrnvqA-IM5e
    - https://medium.com/@leodpereda/create-a-beautiful-3d-map-with-pydeck-geopandas-and-pandas-8cd1d73e1ec3

group Database

endGroup

endGroup

endSection

## Version history

- 0.0.1
  - Initial release
