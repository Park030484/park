from docx import Document

# Creating a new document for the updated performance task
doc_final = Document()

# Add title
doc_final.add_heading('DAILY ACTIVITY LOG', 0)

# Week 1 heading
doc_final.add_heading('Week 1: September 8-14', level=1)

# Add table for week 1
table1_final = doc_final.add_table(rows=8, cols=5)
table1_final.style = 'Table Grid'

# Add headers to the table
headers_final = ["Day", "Warm-up", "Workout", "Cool down", "Picture Placeholder"]
for i, header in enumerate(headers_final):
    table1_final.cell(0, i).text = header

# Week 1 data
week1_data_final = [
    ("Sunday (8)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown, "Insert Picture Here"),
    ("Monday (9)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown, "Insert Picture Here"),
    ("Tuesday (10)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown, "Insert Picture Here"),
    ("Wednesday (11)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown, "Insert Picture Here"),
    ("Thursday (12)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown, "Insert Picture Here"),
    ("Friday (13)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown, "Insert Picture Here"),
    ("Saturday (14)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown, "Insert Picture Here"),
]

# Filling Week 1 data
for i, row_data in enumerate(week1_data_final):
    for j, cell_data in enumerate(row_data):
        table1_final.cell(i+1, j).text = cell_data

# Week 2 heading
doc_final.add_heading('Week 2: September 15-21', level=1)

# Add table for week 2
table2_final = doc_final.add_table(rows=8, cols=5)
table2_final.style = 'Table Grid'

# Add headers to the table
for i, header in enumerate(headers_final):
    table2_final.cell(0, i).text = header

# Week 2 data
week2_data_final = [
    ("Sunday (15)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown, "Insert Picture Here"),
    ("Monday (16)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown, "Insert Picture Here"),
    ("Tuesday (17)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown, "Insert Picture Here"),
    ("Wednesday (18)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown, "Insert Picture Here"),
    ("Thursday (19)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown, "Insert Picture Here"),
    ("Friday (20)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown, "Insert Picture Here"),
    ("Saturday (21)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown, "Insert Picture Here"),
]

# Filling Week 2 data
for i, row_data in enumerate(week2_data_final):
    for j, cell_data in enumerate(row_data):
        table2_final.cell(i+1, j).text = cell_data

# Save the final document
output_final_path = '/mnt/data/Final_Updated_Daily_Activity_Log.docx'
doc_final.save(output_final_path)

output_final_path
