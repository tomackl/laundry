import click
from laundry.constants import laundry_version
from laundry.laundryclass import Laundry
from pathlib import Path



@click.group()
@click.version_option(laundry_version)
def cli():
    """
    This is the command line interface(CLI) for the Laundry app.
    For details regarding the operation of the app type `laundry --help`.
    """
    pass


@cli.command()
@click.option('--data-worksheet', '-dw', 'data',
              default='Master List',
              help='Name of the worksheet containing the cell_data to be converted into a '
                   'word document. '
                   'The default is "Master List".'
              )
@click.option('--template', '-t', 'template',
              help='Name of the template file to be used used as the basis of the '
                   'converted file.',
              type=click.Path(exists=True)
              )
@click.option('--structure-worksheet', '-sw', '-s', 'structure',
              default='_structure',
              help='Name of the worksheet containing the cell_data to format the structure '
                   'of the outfile document. The default is "_structure".'
              )
@click.option('--data-header', '-dh', 'data_head',
              default=0,
              type=int,
              help="The row number of the cell_data worksheet's row containing the column "
                   "headers. The default is 0."
              )
@click.argument('input_file',
                type=click.Path(exists=True)
                )
@click.argument('output_file')
def single(input_file: str, output_file: str, data: str, structure: str, template: str, data_head: int):
    """
    Run laundry on a single worksheet.

    The relative path for each file should be provided with each of the options if non-default file names are provided.

    NOTE: If output files are intended to be saved in a separate directory, that directory *must* exist otherwise the
    output file will not save.

    IMPORTANT: Laundry will overwrite, without prompting, any files with the same name in the directory where output
    files are saved.
    """
    file_input: Path = Path(input_file)
    file_output: str = output_file
    wkst_data: str = data
    wkst_struct: str = structure
    template: str = template
    Laundry(file_input, data_worksheet=wkst_data, structure_worksheet=wkst_struct, template_file=template,
            header_row=data_head, output_file=file_output)


@cli.command()
@click.option('--batch-worksheet', '-b', 'batch',
              default='_batch',
              help='Name of the worksheet containing the format cell_data. This '
                   'worksheet defines the structure and cell_data worksheets and '
                   'other higher level formatting details. The default batch '
                   'worksheet name is "_batch".')
@click.argument('input_file',
                type=click.Path(exists=True)
                )
def multi(input_file: (Path, str), batch: str):
    """
    Run Laundry on multiple worksheets.
    """
    file_input: Path = Path(input_file)
    wksht_batch: str = batch
    Laundry(file_input, batch_worksheet=wksht_batch)
