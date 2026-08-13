lines = open('maas-script.js', encoding='utf-8').readlines()

rep_marker     = '"Rep": "Michael Shonde"'
email_marker   = '"Rep Email": "Michael.Shonde@ibm.com"'

rep_new        = '"Rep": "Marianne Johnston"'
email_new      = '"Rep Email": "mjohnston@ibm.com"'

changed = 0
for idx, line in enumerate(lines):
    if rep_marker in line:
        # Check next 6 lines to confirm this is a Marianne record
        chunk = ''.join(lines[idx:idx+6])
        if 'Marianne Johnston' in chunk:
            lines[idx] = line.replace(rep_marker, rep_new)
            changed += 1
    if email_marker in line:
        chunk_back = ''.join(lines[max(0,idx-3):idx+3])
        if 'Marianne Johnston' in chunk_back:
            lines[idx] = line.replace(email_marker, email_new)

with open('maas-script.js', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print(f'Updated {changed} records.')
