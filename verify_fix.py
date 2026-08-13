lines = open('maas-script.js', encoding='utf-8').readlines()
remaining  = [i+1 for i,l in enumerate(lines) if 'Michael Shonde' in l]
mar_as_rep = [i+1 for i,l in enumerate(lines) if '"Rep": "Marianne Johnston"' in l]
print('Michael Shonde remaining:', len(remaining), '(should be 4 - Josh Knowles records)')
print('Marianne Johnston as Rep:', len(mar_as_rep), '(should be 16)')
for ln in remaining:
    print(' line', ln, ':', lines[ln-1].strip())
