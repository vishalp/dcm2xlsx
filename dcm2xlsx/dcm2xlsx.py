#!/usr/bin/env python3

import argparse
import pathlib
import re
import sys
import pandas as pd
import pydicom
import xlsxwriter

def convert(dicom_directory, logger=None):
    if logger is None:
        logger = lambda x: print(x, file=sys.stderr)
    logger(f'Scanning {dicom_directory} ...')
    all_files = sorted(dicom_directory.glob('*'))
    dicom_files = []
    tags = ['PatientID', 'PatientName', 'StudyInstanceUID', 'StudyDescription', 'SeriesInstanceUID', 'SeriesDescription']
    df = []
    for f in all_files:
        try:
            ds = pydicom.dcmread(f, stop_before_pixels=True, specific_tags=tags)
            df.append(tuple([ds.get(t) for t in tags]))
            dicom_files.append(f)
        except:
            continue
    df = pd.DataFrame.from_records(df, columns=tags)
    logger(f'Found {len(dicom_files)} DICOM files ...')

    bad_chars = re.compile(r'[^\w.-]')
    path_str = lambda x, y: re.sub(bad_chars, '_', str(x)) if x is not None else y
    cell_size = 16
    for patient in df['PatientID'].unique():
        df_patient = df[df['PatientID'] == patient]
        patient_dir = dicom_directory / path_str(df_patient['PatientName'].iloc[0], 'UnknownPatient')
        for study in df_patient['StudyInstanceUID'].unique():
            df_study = df_patient[df_patient['StudyInstanceUID'] == study]
            study_dir = patient_dir / path_str(df_study['StudyDescription'].iloc[0], 'UnknownStudy')
            study_dir.mkdir(parents=True, exist_ok=True)
            for series in df_study['SeriesInstanceUID'].unique():
                df_series = df_study[df_study['SeriesInstanceUID'] == series]
                series_file = study_dir / path_str(df_series['SeriesDescription'].iloc[0], 'UnknownSeries')
                series_file = series_file.with_suffix('.xlsx')

                book = xlsxwriter.Workbook(series_file)
                for i in df_series.index:
                    logger(f'Converting {dicom_files[i]} ...')
                    ds = pydicom.dcmread(dicom_files[i])
                    sheet_name = ds.get('InstanceNumber')
                    sheet_name = f'Instance{sheet_name}' if sheet_name is not None else ''
                    sheet = book.add_worksheet(sheet_name)
                    arr = ds.pixel_array
                    n_rows, n_cols = arr.shape
                    for row in range(n_rows):
                        for col in range(n_cols):
                            sheet.write_number(row, col, arr[row, col])
                    wc = ds.get('WindowCenter')
                    ww = ds.get('WindowWidth')
                    ps = ds.get('PixelSpacing')
                    if wc is not None and ww is not None and ps is not None:
                        sheet.conditional_format(0, 0, n_rows - 1, n_cols - 1, {
                            'type': '2_color_scale', 
                            'min_type': 'num', 'min_value': wc - ww / 2, 'min_color': '#000000', 
                            'max_type': 'num', 'max_value': wc + ww / 2, 'max_color': '#FFFFFF'
                        })
                        for row in range(n_rows):
                            sheet.set_row_pixels(row, cell_size)
                        sheet.set_column_pixels(0, n_cols - 1, cell_size * ps[1] / ps[0])
                logger(f'Writing {series_file} ...')
                book.close()
    logger('Done.')

def main():
    def is_existing_dir(s):
        p = pathlib.Path(s)
        if not p.is_dir():
            raise argparse.ArgumentTypeError(f'{s} is not an existing directory!')
        return p.absolute()
    parser = argparse.ArgumentParser(description='Convert DICOM files to OOXML spreadsheets.')
    parser.add_argument('dicom_directory', type=is_existing_dir, help='directory containing DICOM images')
    args = parser.parse_args()
    convert(args.dicom_directory)

if __name__ == '__main__':
    main()
