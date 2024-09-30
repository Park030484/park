# park
# Reattempting with correct row indexing, manually adding content without looping errors.

# Create a new document again
doc_new_retry = Document()

# Add title
doc_new_retry.add_heading('DAILY ACTIVITY LOG', 0)

# Week 1 heading
doc_new_retry.add_heading('Week 1: September 8-14', level=1)

# Add table for week 1
table1_retry = doc_new_retry.add_table(rows=8, cols=4)
table1_retry.style = 'Table Grid'

# Add headers to the table
headers_retry = ["Day", "Warm-up", "Workout", "Cool down"]
for i, header in enumerate(headers_retry):
    table1_retry.cell(0, i).text = header

# Manually add week 1 data
week1_data = [
    ("Sunday (8)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown),
    ("Monday (9)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown),
    ("Tuesday (10)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown),
    ("Wednesday (11)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown),
    ("Thursday (12)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown),
    ("Friday (13)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown),
    ("Saturday (14)", week1_warmup, f"Circuit 1: {week1_circuit1}\nCircuit 2: {week1_circuit2}\nRest 2-3 min", week1_cooldown),
]

for i, row_data in enumerate(week1_data):
    for j, cell_data in enumerate(row_data):
        table1_retry.cell(i+1, j).text = cell_data

# Week 2 heading
doc_new_retry.add_heading('Week 2: September 15-21', level=1)

# Add table for week 2
table2_retry = doc_new_retry.add_table(rows=8, cols=4)
table2_retry.style = 'Table Grid'

# Add headers to the table
for i, header in enumerate(headers_retry):
    table2_retry.cell(0, i).text = header

# Manually add week 2 data
week2_data = [
    ("Sunday (15)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown),
    ("Monday (16)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown),
    ("Tuesday (17)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown),
    ("Wednesday (18)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown),
    ("Thursday (19)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown),
    ("Friday (20)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown),
    ("Saturday (21)", week2_warmup, f"Circuit 1: {week2_circuit1}\nCircuit 2: {week2_circuit2}\nRest 2-3 min", week2_cooldown),
]

for i, row_data in enumerate(week2_data):
    for j, cell_data in enumerate(row_data):
        table2_retry.cell(i+1, j).text = cell_data

# Save the document
output_retry_path = '/mnt/data/Updated_Daily_Activity_Log_Final.docx'
doc_new_retry.save(output_retry_path)

output_retry_path
