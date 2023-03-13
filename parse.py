import pathlib
import openpyxl
import json

def parse_hex_value(s):
  if type(s) != str: s = str(int(s))
  return int(s, 16)
def prepare_directory(path):
  path = pathlib.Path(path)
  path.parent.mkdir(parents=True, exist_ok=True)

'''
r13db = {
  'type': None,
  'GMSJ01': 0x80410AC0,
  'GMSE01': 0x804141C0,
  'GMSP01': 0x8040B960,
  'GMSJ0A': 0x804051A0,
}
'''
def parse_static_variables(wb, out_path):
  ws = wb['Static variables']
  itr_row = ws.rows
  db = {}
  ## Row 1
  row1 = next(itr_row)
  colidx = {c.value: ic for ic, c in enumerate(row1)}
  ic_name = colidx['Name']
  ic_type = colidx['Type']
  ic_addr = colidx['Absolute address']
  ## Row 2
  row2 = next(itr_row)
  versions = ['GMSJ01', 'GMSE01', 'GMSP01', 'GMSJ0A'] # FIXME read from file
  ## entries
  for row in itr_row:
    db[row[ic_name].value] = {
      'type': row[ic_type].value,
      **{
        ver: parse_hex_value(row[ic].value)
        for ic, ver in enumerate(versions, start=ic_addr)
      },
    }
  ## write
  prepare_directory(out_path)
  with open(out_path, 'w') as f:
    json.dump(db, f)

if __name__ == '__main__':
  wb = openpyxl.open('SMS RAM Map.xlsx')
  parse_static_variables(wb, 'docs/static-variables.json')
