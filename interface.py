from doc_edit import DocEdit

# Init the program
doc_edit = DocEdit()
document = doc_edit.open_document('C:/Users/LENOVO/Desktop/Contrato')

# Create a form
# doc_edit.create_form(
#     'C:/Users/LENOVO/Desktop/formtest',
#     'nome', 'idade', 'cpf'
# )

# Upload created form
form = doc_edit.upload_form(
    'C:/Users/LENOVO/Desktop/formtest',
)

# Updating the uploaded form
for key, value in form.items():
    value = str(input(f'{key}: '))
    form[key] = value

# Using the uploaded form to fill the document fields
doc_edit.update_document(
    document,
    'C:/Users/LENOVO/Desktop/Contrato2',
    form
)
