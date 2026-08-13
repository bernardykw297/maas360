import re

lines = open('maas-script.js', encoding='utf-8').readlines()
shonde_idxs = [i for i, l in enumerate(lines) if '"Rep": "Michael Shonde"' in l]
email_idxs  = [i for i, l in enumerate(lines) if '"Rep Email": "Michael.Shonde@ibm.com"' in l]

# All 20 records are NY -> Steven St. Arnauld
REP_OLD   = '"Rep": "Michael Shonde"'
REP_NEW   = '"Rep": "Steven St. Arnauld"'
EMAIL_OLD = '"Rep Email": "Michael.Shonde@ibm.com"'
EMAIL_NEW = '"Rep Email": "steven.st.Arnauld@ibm.com"'

changed_rep   = 0
changed_email = 0

for i in shonde_idxs:
    lines[i] = lines[i].replace(REP_OLD, REP_NEW)
    changed_rep += 1

for i in email_idxs:
    lines[i] = lines[i].replace(EMAIL_OLD, EMAIL_NEW)
    changed_email += 1

with open('maas-script.js', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print(f'Rep updated:   {changed_rep}')
print(f'Email updated: {changed_email}')

# Verify nothing left
remaining = [i+1 for i, l in enumerate(lines) if 'Michael Shonde' in l]
print(f'Michael Shonde remaining: {len(remaining)} (should be 0)')
