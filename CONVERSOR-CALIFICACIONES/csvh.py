"""
This file is part of ToR Conversor.

    ToR Conversor is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    ToR Conversor is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with ToR Conversor.  If not, see <https://www.gnu.org/licenses/>.
"""

import csv


def addHeader(csv_data, header):
    """ Add a header to the CSV data """
    csv_data.insert(0, header)
    return csv_data

def importRawCSVData(file_name, delimiter=";", quotechar='"', remove_empty_cells=False):
    """ Import a CSV file as a 2D array """
    data = []
    with open(file_name, 'r',encoding='utf-8-sig') as csvfile:

        reader = csv.reader(csvfile,
                            delimiter=delimiter,
                            quotechar=quotechar)

        for row in reader:

            data.append([])
            for cell in row:
                d = cell.strip()
                v = d


                if remove_empty_cells and cell.strip() != "":
                    data[-1].append(v)
                elif not remove_empty_cells:

                    data[-1].append(v)

    return data


def exportRawCSVData(file_name, csv_data, delimiter=";", quotechar='"', add=False):
    """ Export a CSV file as a 2D array """
    if add:
        mode = 'a+'
    else:
        mode = 'w'

    with open(file_name, mode, newline='') as csvfile:
        writer = csv.writer(csvfile,
                            delimiter=delimiter,
                            quotechar=quotechar,
                            quoting=csv.QUOTE_MINIMAL)
        for row in csv_data:
            writer.writerow(row)
