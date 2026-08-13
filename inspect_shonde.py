import re

lines = open('maas-script.js', encoding='utf-8').readlines()
shonde_idxs = [i for i, l in enumerate(lines) if '"Rep": "Michael Shonde"' in l]

print(f'Total: {len(shonde_idxs)} records\n')
for idx in shonde_idxs:
    # grab surrounding context
    chunk = ''.join(lines[max(0,idx-15):idx+8])
    coverage  = re.search(r'"COVERAGE_NAME":\s*"([^"]+)"', chunk)
    branch    = re.search(r'"BRANCH_NAME":\s*"([^"]+)"', chunk)
    branch_unit = re.search(r'"BRANCH_UNIT_NAME":\s*"([^"]+)"', chunk)
    tech_rep  = re.search(r'"Tech Rep":\s*"([^"]+)"', chunk)
    print(f'  line {idx+1}')
    print(f'    Coverage   : {coverage.group(1) if coverage else "?"}')
    print(f'    Branch     : {branch.group(1) if branch else "?"}')
    print(f'    Branch Unit: {branch_unit.group(1) if branch_unit else "?"}')
    print(f'    Tech Rep   : {tech_rep.group(1) if tech_rep else "?"}')
    print()
