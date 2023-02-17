import json
from docx import Document
from docx.opc.exceptions import PackageNotFoundError


class EditDoc:
    def __init__(self) -> None:
        self.form = None

    def open_document(self, dir: str) -> Document:
        """
        Open the document .docx to be edited.

        :param dir: document's directory
        :return .docx file
        """
        try:
            document = Document(dir)
        except PackageNotFoundError as error:
            return print(f'{error}')

        return document

    def upload_form(self, dir: str, *args):
        """
        Upload a .json form with already saved fields that
        can be used as a parameter in update_document() function.

        :param *args: values for form keys
        :param dir: form's directory
        :return dict
        """
        try:
            with open(dir, 'r') as fp:
                data = json.load(fp)

                if len(data) >= len(args):
                    for i, field in enumerate(data):
                        data[field] = args[i]
                else:
                    return None

                return data

        except FileNotFoundError as error:
            return print(f'{error}')

    def save_form(self, save_dir):
        """
        Saves the keys used in the update_document() function in
        a .json file to be reused as a parameter for the same function.

        :param save_dir: file's save directory
        """
        try:
            form = self.form
            save = save_dir

            for i in form:
                form[i] = ''

            with open(save, 'w') as fp:
                json.dump(form, fp)

        except (FileNotFoundError, PermissionError, TypeError) as error:
            return print(f'{error}')

    def update_document(self, document, save_dir, form=None, **kwargs):
        """
        Update the values of the keys to be changed in the document .docx;
        If you have a .json file with predefined keys, you can use it here,
        passing it as an argument to "form", or pass the keys and values
        trough the **kwargs parameter.

        :param document: .docx file
        :param save_dir: file's save directory
        :param form: A form of parameters fields
        :param **kwargs: Key and values parameters
        """
        try:
            if form is not None:
                fields = form
            else:
                fields = {}
                fields.update(kwargs)
                fields = fields['kwargs']

            for paragraph in document.paragraphs:
                for field in fields:
                    value = fields[field]
                    paragraph.text = paragraph.text.replace(
                        field, str(value)
                    )

            document.save(f'{save_dir}.docx')
            self.form = fields

        except (FileNotFoundError, KeyError) as error:
            return print(f'{error}')
