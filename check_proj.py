import re
code = open("Automat_andre.py", encoding="utf-8").read()
cs = set(re.findall(r'\("(\d+A{3})", "(\d+Z{3})"', code))
lines = [l.strip() for l in open("Projetos.txt", encoding="utf-8") if l.strip()][1:]
ps = set((l.split()[2], l.split()[3]) for l in lines)
print("codigo:", len(cs), "| projetos.txt:", len(ps))
print("So no codigo:", sorted(cs - ps))
print("So no Projetos.txt:", sorted(ps - cs))
