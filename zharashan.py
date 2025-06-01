from datetime import datetime
import os
from python_docx import Document
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


def generate_document(
        input_fullname,
        input_date=None,
        output_path='output.docx',
        template_path=None,
        debug=False
):
    """
    Generate a Word document with the given inputs

    Args:
        input_fullname (str): Full name of the person
        input_date (datetime, optional): Date to insert. Defaults to current date
        output_path (str): Path where to save the generated document
        template_path (str, optional): Custom path to template file
        debug (bool): Enable debug mode
    """
    if debug:
        logging.getLogger().setLevel(logging.DEBUG)

    # Verify template existence
    if template_path is None:
        template_path = os.path.join(os.path.dirname(__file__), 'templates', 'template.docx')

    if not os.path.exists(template_path):
        logger.error(f"Template file not found at: {template_path}")
        return False

    # Format the date
    if input_date is None:
        input_date = datetime.now()

    month_replacements = {
        'January': 'январь',
        'February': 'февраль',
        'March': 'март',
        'April': 'апрель',
        'May': 'май',
        'June': 'июнь',
        'July': 'июль',
        'August': 'август',
        'September': 'сентябрь',
        'October': 'октябрь',
        'November': 'ноябрь',
        'December': 'декабрь'
    }

    formatted_date = input_date.strftime('«%d» %B %Y')
    for eng, rus in month_replacements.items():
        formatted_date = formatted_date.replace(eng, rus)

    logger.info(f"Formatted date: {formatted_date}")
    logger.info(f"Full name: {input_fullname}")

    try:
        # Open the document
        doc = Document(template_path)

        # Replace in paragraphs
        for paragraph in doc.paragraphs:
            if '{{ input_date }}' in paragraph.text:
                logger.info(f"Found date placeholder in paragraph: {paragraph.text}")
                paragraph.text = paragraph.text.replace('{{ input_date }}', formatted_date)

            if '{{ input_fullname }}' in paragraph.text:
                logger.info(f"Found name placeholder in paragraph: {paragraph.text}")
                paragraph.text = paragraph.text.replace('{{ input_fullname }}', input_fullname)

        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if '{{ input_date }}' in paragraph.text:
                            logger.info(f"Found date placeholder in table cell: {paragraph.text}")
                            paragraph.text = paragraph.text.replace('{{ input_date }}', formatted_date)

                        if '{{ input_fullname }}' in paragraph.text:
                            logger.info(f"Found name placeholder in table cell: {paragraph.text}")
                            paragraph.text = paragraph.text.replace('{{ input_fullname }}', input_fullname)

        # Save the document
        doc.save(output_path)
        logger.info(f"Document generated successfully at: {output_path}")
        return True

    except Exception as e:
        logger.error(f"Error generating document: {e}")
        return False


if __name__ == "__main__":
    # Example usage
    current_dir = os.path.dirname(__file__)
    template_path = os.path.join(current_dir, 'templates', 'template.docx')

    sample_data = {
        "input_fullname": "Иванов Иван Иванович",
        "input_date": datetime.now(),
        "template_path": template_path,
        "debug": True
    }

    success = generate_document(**sample_data)
    if success:
        print("Document generated successfully!")
    else:
        print("Error generating document. Check logs for details.")
