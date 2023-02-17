import json
from docx import Document
from docx.opc.exceptions import PackageNotFoundError


class EditDoc:
    def __init__(self) -> None:
        self.form = None

    def open_document(self, dir: str) -> Document:
        """
        Open the document to be edited.

        :param dir: document's directory
        :return .docx file
        """
        try:
            document = Document(dir)
        except PackageNotFoundError as error:
            return print(f'{error}')

        return document

    def update_document(self, document, save_dir, form=None, **kwargs):
        """
        Update the document fields to new chosen ones.

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

    def upload_form(self, dir: str, *args):
        """
        Upload a json form with already saved fields

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
        Save the used fields into a json form

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
