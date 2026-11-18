import json,re
def js(o):return json.dumps(o,ensure_ascii=False,separators=(',',':')).replace('</','<\\/')
KEYS=('num','title','page','levels','paths','options','kind','text')
man=[{k:s[k] for k in KEYS} for s in json.load(open('manual.json'))]
figs=json.load(open('figs.json'));figs['felt']['cap']='Illustration 73: Changing the felt strip in the printer';figs['device']['cap']='Illustration 1: Device view GLM-I'
prods=json.load(open('prices.json'))
tpl=open('template.html').read()
style=re.search(r'<style>([\s\S]*?)</style>',tpl).group(1)+open('v6.css').read()
body=re.search(r'<body>([\s\S]*)$',tpl).group(1)
main=open('app_main.js').read();boot=open('app_boot.js').read()
b2=''.join(open(f).read() for f in ['b2a.js','b2b.js','b2c.js','b2d.js','dm.js','b2e.js'])
app=main+'\n'+b2+'\n'+boot
open('app_full.js','w').write(app)
data='const MAN='+js(man)+';\nconst SITE_PRODUCTS='+js(prods)+';\nconst DEMO_PRODUCTS=SITE_PRODUCTS;\nconst FIGS='+js(figs)+';\n'
head='<title>Labeller Line Trainer</title>\n<meta name="claude:viewport-fit-sides" content="cover">\n'
art=head+'<style>'+style+'</style>\n'+body+'\n<script>\n'+data+'</script>\n<script>\n'+app+'\n</script>\n'
open('labeller.html','w').write(art)
full='<!doctype html>\n<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'+head+'<style>'+style+'</style></head><body>'+body+'\n<script>\n'+data+'</script>\n<script>\n'+app+'\n</script>\n</body></html>\n'
open('Labeller Line Trainer V6.html','w').write(full)
print(len(art)//1024,'KB')
