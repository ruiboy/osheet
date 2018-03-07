#!/usr/bin/env python

import argparse, csv, os, sys, xlsxwriter

# constants
expected_csv_header = \
  ['Date', ' "Session"', '', ' "Child"', ' "Parent"', ' "Email"', ' "Emergency contact phone number"', ' "Allergies"']

i_date = 0
i_session = 1
i_kid = 3
i_allergies = 7

am_session = 'Before school care'
pm_session = 'After school care'
allergies_false_positive = ['none', 'nil', 'no', 'nothing', 'n/a']

expected_allergy_conf_header = ['Child', 'Tags']
expected_skip_conf_header = ['Child', 'When']

extra_signin_rows = 25
total_booking_rows = 75
total_booking_rows = 75

year = 2018
term = 1
week = 6

# working
allergy_conf = []
skip_conf = []
debug = False


#######################
# Signin Sheets
#######################

def make_signin_sheets(in_csv_file, out_dir):
  with open(in_csv_file, 'rb') as in_file:
    data = prepare_signin_data(new_csv_reader(in_file), am_session)
    write_am_signin_sheet(out_dir, data)

  with open(in_csv_file, 'rb') as in_file:
    data = prepare_signin_data(new_csv_reader(in_file), pm_session)
    write_pm_signin_sheet(out_dir, data)


def prepare_signin_data(reader, session):
  # list of kid -> kid
  data = {}

  for row in reader:
    if len(row) == len(expected_csv_header) and tidy(row[i_session]) == session:
      kid = flip_name(tidy(row[i_kid]))
      if not skip_kid(kid) and not kid in data:
        data[kid] = kid

  if debug: print 'Signin data: %s: %s' % (session, data)
  return data


def write_am_signin_sheet(out_dir, data):
  out_file = os.path.join(out_dir, 'am-signin.xlsx')
  wb = xlsxwriter.Workbook(out_file);
  ws = wb.add_worksheet()
  formats = make_signin_formats(wb)

  print 'Writing: %s' % out_file

  ws.set_column(0, 0, 20)
  ws.set_column(1, 16, 5)

  ws.set_row(0, 22)
  ws.write('B1', 'AM Attendance %s' % year, formats['h1'])

  ws.set_row(1, 16)
  ws.write('A2', 'Week: %s' % week, formats['h2'])
  ws.write('B2', 'Term: %s' % term, formats['h2'])

  ws.set_row(3, 16)
  ws.write('A4', '', formats['th1'])
  ws.merge_range('B4:D4', 'Monday', formats['th1'])
  ws.merge_range('E4:G4', 'Tuesday', formats['th1'])
  ws.merge_range('H4:J4', 'Wednesday', formats['th1'])
  ws.merge_range('K4:M4', 'Thursday', formats['th1'])
  ws.merge_range('N4:P4', 'Friday', formats['th1'])

  ws.set_row(4, 16)
  ws.write('A5', 'Name', formats['th2Name'])
  for col in [1, 4, 7, 10, 13]:
    ws.write(4, col, 'Time in', formats['th2Left'])
    ws.write(4, col + 1, 'Sign', formats['th2Middle'])
    ws.write(4, col + 2, 'Mark Off', formats['th2Right'])

  write_signin_sheet_data(ws, formats, data)

  wb.close()


def write_pm_signin_sheet(out_dir, data):
  out_file = os.path.join(out_dir, 'pm-signin.xlsx')
  wb = xlsxwriter.Workbook(out_file);
  ws = wb.add_worksheet()
  formats = make_signin_formats(wb)

  print 'Writing: %s' % out_file

  ws.set_column(0, 0, 20)
  ws.set_column(1, 16, 5)

  ws.set_row(0, 22)
  ws.write('B1', 'PM Attendance %s' % year, formats['h1'])

  ws.set_row(1, 16)
  ws.write('A2', 'Week: %s' % week, formats['h2'])
  ws.write('B2', 'Term: %s' % term, formats['h2'])

  ws.set_row(3, 16)
  ws.write('A4', '', formats['th1'])
  ws.merge_range('B4:D4', 'Monday', formats['th1'])
  ws.merge_range('E4:G4', 'Tuesday', formats['th1'])
  ws.merge_range('H4:J4', 'Wednesday', formats['th1'])
  ws.merge_range('K4:M4', 'Thursday', formats['th1'])
  ws.merge_range('N4:P4', 'Friday', formats['th1'])

  ws.set_row(4, 16)
  ws.write('A5', 'Name', formats['th2Name'])
  for col in [1, 4, 7, 10, 13]:
    ws.write(4, col, 'Time in', formats['th2Left'])
    ws.write(4, col + 1, 'Time out', formats['th2Middle'])
    ws.write(4, col + 2, 'Sign', formats['th2Right'])

  write_signin_sheet_data(ws, formats, data)

  wb.close()


def write_signin_sheet_data(ws, formats, data):
  suffix = 'Top'
  row = 5
  for kid in sorted(data):
    write_signin_sheet_row(ws, formats, suffix, row, kid)
    row += 1
    suffix = ''

  for blah in range(0, extra_signin_rows):
    write_signin_sheet_row(ws, formats, suffix, row, '')
    row += 1
    suffix = ''

  write_signin_sheet_row(ws, formats, 'Underneath', row, '')


def write_signin_sheet_row(ws, formats, format_suffix, row, kid):
  ws.write(row, 0, kid, formats['tdName' + format_suffix])
  for col in [1, 4, 7, 10, 13]:
    ws.write(row, col, '', formats['tdLeft' + format_suffix])
    ws.write(row, col + 1, '', formats['tdMiddle' + format_suffix])
    ws.write(row, col + 2, '', formats['tdRight' + format_suffix])


def make_signin_formats(wb):
  return {
    'h1': wb.add_format({'font_size': 20, 'bold': True}),
    'h2': wb.add_format({'font_size': 14, 'bold': True}),

    'th1': wb.add_format(
      {'font_size': 14, 'bold': True, 'top': 2, 'right': 2, 'bottom': 2, 'left': 2, 'align': 'center'}),

    'th2Name': wb.add_format({'font_size': 14, 'bold': True, 'top': 2, 'right': 2, 'bottom': 2, 'left': 2}),
    'th2Left': wb.add_format({'font_size': 8, 'top': 2, 'right': 1, 'bottom': 2, 'left': 2, 'align': 'center'}),
    'th2Middle': wb.add_format({'font_size': 8, 'top': 2, 'right': 1, 'bottom': 2, 'left': 1, 'align': 'center'}),
    'th2Right': wb.add_format({'font_size': 8, 'top': 2, 'right': 2, 'bottom': 2, 'left': 1, 'align': 'center'}),

    'tdNameTop': wb.add_format({'font_size': 10, 'top': 2, 'right': 2, 'bottom': 1, 'left': 2}),
    'tdLeftTop': wb.add_format({'font_size': 10, 'top': 2, 'right': 1, 'bottom': 1, 'left': 2}),
    'tdMiddleTop': wb.add_format({'font_size': 10, 'top': 2, 'right': 1, 'bottom': 1, 'left': 1}),
    'tdRightTop': wb.add_format({'font_size': 10, 'top': 2, 'right': 2, 'bottom': 1, 'left': 1}),

    'tdName': wb.add_format({'font_size': 10, 'top': 1, 'right': 2, 'bottom': 1, 'left': 2}),
    'tdLeft': wb.add_format({'font_size': 10, 'top': 1, 'right': 1, 'bottom': 1, 'left': 2}),
    'tdMiddle': wb.add_format({'font_size': 10, 'top': 1, 'right': 1, 'bottom': 1, 'left': 1}),
    'tdRight': wb.add_format({'font_size': 10, 'top': 1, 'right': 2, 'bottom': 1, 'left': 1}),

    'tdNameUnderneath': wb.add_format({'font_size': 10, 'top': 2}),
    'tdLeftUnderneath': wb.add_format({'font_size': 10, 'top': 2}),
    'tdMiddleUnderneath': wb.add_format({'font_size': 10, 'top': 2}),
    'tdRightUnderneath': wb.add_format({'font_size': 10, 'top': 2}),
  }


#######################
# Booking Sheets
#######################

def make_booking_sheets(in_csv_file, out_dir):
  with open(in_csv_file, 'rb') as in_file:
    data = prepare_booking_data(new_csv_reader(in_file), am_session)
    write_booking_sheet(out_dir, 'am-booking.xlsx', 'AM', data)

  with open(in_csv_file, 'rb') as in_file:
    data = prepare_booking_data(new_csv_reader(in_file), pm_session)
    write_booking_sheet(out_dir, 'pm-booking.xlsx', 'PM', data)


def prepare_booking_data(reader, session):
  # dict of dicts: date -> kid -> allergies
  data = {}

  for row in reader:
    if len(row) == len(expected_csv_header) and tidy(row[i_session]) == session:
      kid = flip_name(tidy(row[i_kid]))

      if not skip_kid(kid):
        date = tidy(row[i_date])
        allergies = tidy(row[i_allergies])

        if not date in data:
          data[date] = {}

        kids = data[date]
        if not kid in kids:
          kids[kid] = allergies

  if debug: print 'Booking data: %s: %s' % (session, data)
  return data


def write_booking_sheet(out_dir, out_file_name, label, data):
  out_file = os.path.join(out_dir, out_file_name)
  wb = xlsxwriter.Workbook(out_file);
  ws = wb.add_worksheet()
  formats = make_booking_formats(wb)

  print 'Writing: %s' % out_file

  # set col widths
  small, big = 5, 20
  for col in range(0, 11):
    ws.set_column(col, col, big if col % 2 else small)

  # write headings
  ws.set_row(0, 22)
  ws.merge_range('A1:K1', '%s Bookings' % label, formats['h1'])

  ws.set_row(1, 26)
  for col in range(0, 11):
    ws.write(1, col, '', formats['th1'])
  ws.write('B2', 'Week %s' % week, formats['th1'])

  for col in range(0, 11):
    ws.write(2, col, '', formats['th2'])

  # format a bunch of blank rows
  for row in range(1, total_booking_rows):
    ws.write(row + 2, 0, row, formats['tdNum'])
    for col in range(1, 11):
      ws.write(row + 2, col, '', formats['td'])

  # write in the data
  col = 1
  for date, kids in sorted(data.iteritems()):
    ws.write(2, col, date, formats['th2'])
    row = 3
    for kid, allergies in sorted(kids.iteritems()):
      ws.write(row, col, kid, formats['tdAllergies' if has_allergies(kid, allergies) else 'td'])
      tag = get_allergy_tag(kid)
      if tag:
        ws.write(row, col + 1, tag, formats['td'])
      row += 1
    col += 2

  wb.close()


def make_booking_formats(wb):
  return {
    'h1': wb.add_format({'font_size': 20, 'bold': True, 'align': 'center'}),

    'th1': wb.add_format(
      {'font_size': 24, 'bold': True, 'top': 2, 'right': 2, 'bottom': 2, 'left': 2, 'align': 'center'}),
    'th2': wb.add_format(
      {'font_size': 12, 'bold': True, 'top': 2, 'right': 2, 'bottom': 2, 'left': 2, 'align': 'center',
       'bg_color': '#CCCCCC'}),

    'tdNum': wb.add_format(
      {'font_size': 12, 'bold': True, 'top': 2, 'right': 2, 'bottom': 2, 'left': 2, 'align': 'center'}),
    'td': wb.add_format(
      {'font_size': 10, 'top': 2, 'right': 2, 'bottom': 2, 'left': 2}),
    'tdAllergies': wb.add_format(
      {'font_size': 10, 'top': 2, 'right': 2, 'bottom': 2, 'left': 2, 'bg_color': '#00CC00'}),
  }


#######################
# Helpers
#######################

def assert_csv(in_csv_file):
  if not os.path.isfile(in_csv_file):
    error('csv file does not exist: %s' % in_csv_file)

  print 'Reading: %s' % in_csv_file

  with open(in_csv_file, 'rb') as in_file:
    reader = new_csv_reader(in_file)
    for row in reader:
      if row != expected_csv_header:
        error('Unexpected csv header.\n Expected:%s\n Got:     %s' % (expected_csv_header, row))
      return


def new_csv_reader(in_file):
  return csv.reader(in_file, delimiter=',', quotechar='"', skipinitialspace=True)


def load_conf(conf_file, expected_header):
  # dict of row[0] -> row
  conf = {}
  print 'Using conf: %s' % conf_file
  if os.path.isfile(conf_file):
    with open(conf_file, 'rb') as in_file:
      reader = csv.reader(in_file, delimiter=',', quotechar='"', skipinitialspace=True)
      header = True
      for row in reader:
        if header:
          if row != expected_header:
            error('Unexpected config header.\n Expected:%s\n Got:     %s' % (expected_header, row))
          header = False
        else:
          conf[row[0]] = row
  if debug: print 'Conf: %s' % conf
  return conf


def skip_kid(kid):
  return kid in skip_conf


def has_allergies(kid, allergies):
  return (allergies and not allergies.lower() in allergies_false_positive) or kid in allergy_conf


def get_allergy_tag(kid):
  return allergy_conf[kid][1] if kid in allergy_conf else ''


def tidy(field):
  return field.strip().strip('"')


def flip_name(name):
  # First Middle Last -> Last, First Middle
  tokens = name.split()
  return tokens[-1] + ', ' + ' '.join(tokens[:-1])


def error(msg):
  print 'ERROR: ' + msg
  sys.exit(1)


def main():
  parser = argparse.ArgumentParser(description='Convert OSH care booking csv to sign-in + booking sheets',
                                   formatter_class=argparse.ArgumentDefaultsHelpFormatter)
  parser.add_argument('--in', help='care booking csv to process', default='in/in.csv')
  parser.add_argument('--out', help='dir for output files', default='out')
  parser.add_argument('--allergy_file', help='csv file of kids and their allergy tags', default='conf/allergy.csv')
  parser.add_argument('--skip_file', help='csv file of kids to ignore', default='conf/skip.csv')
  parser.add_argument('--debug', help='spew out some debug', action='store_true', default=False)
  args = vars(parser.parse_args())

  working_dir = os.path.dirname(os.path.abspath(__file__))

  csv_file = args['in']
  out_dir = args['out']
  global debug
  debug = args['debug']

  assert_csv(csv_file)

  if os.path.isfile(out_dir):
    error('out dir is an existing file; cowardly refusing to do any work')
  if not os.path.isdir(out_dir):
    os.makedirs(out_dir)

  global allergy_conf, skip_conf
  allergy_conf = load_conf(args['allergy_file'], expected_allergy_conf_header)
  skip_conf = load_conf(args['skip_file'], expected_skip_conf_header)

  make_signin_sheets(csv_file, out_dir)
  make_booking_sheets(csv_file, out_dir)


if __name__ == "__main__":
  main()
