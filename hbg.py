# gerekli kütüphaneleri dahil ettim
import random,xlsxwriter

#rasgele renk
r = lambda: random.randint(0,255)

#i,j ve k'yı kullanıcıdan alıp, i ve j'yı liste çevirdim. k'nında sayı olduğundan emin oldum.
i=input("I:")
j=input("J:")
k=int(input("k:"))
i=list(i)
j=list(j)
fasta=[]

# döngülerde “list index out of range” hatasını yaşamamak için i ve j yi k kadar boş elemanlar ile uzattım.
for a in range(0,k):
    i.append(" ")
    j.append(" ")

# ixj boyutlu bir matris yarattım
for a in range(0,len(i)):
    fasta.append([])
    for b in range(0,len(j)):
        fasta[a].append(" ")

#fasta matrisinin içini kurallara uygun şekilde doldurdum
for a in range(0,len(i)-k):
    for b in range(0,len(j)-k):
        #eşleşme olduğunda t sayacını başlatıyorum ve k kadar sağ aşağıyı kontrol ediyorum. k ve t eşit ise sağ aşağı doğru k kadar yıldız atıyorum
        if i[a]==j[b]:
            t=0
            for c in range(0,k):
                if i[a+c]==j[b+c]:
                    t+=1
            if t>=k:
                for c in range(0,t):
                    if a<len(i)-k-2 and b<len(j)-k-2:
                        fasta[a+c][b+c]="*"
                    
#xlsxwriter kütüphanesi ile word dosyası oluşturma işlemleri
workbook = xlsxwriter.Workbook('Desktop\hbg.xlsx')
worksheet = workbook.add_worksheet()

#öncelikle kullanıcıdan alınan i ve j yi yazdırmak için yazı formatı ayarlıyorum
baslik = workbook.add_format({'bold': True,'font_size':16,'align':'center'})
#burada i ve j yazıyorum ardından aşağıdaki for döngüleri ile yanlarına i ve j yi dolduruyorum
worksheet.write(0, 0," ",baslik)
worksheet.write(0, 1,"j",baslik)
worksheet.write(1, 0,"i",baslik)
row=2
col=2
for a in j:
    worksheet.write(0, col,col-2,baslik)
    worksheet.write(1, col,a,baslik)
    col+=1

for a in i:
    worksheet.write(row, 0,row-2,baslik)
    worksheet.write(row, 1,a,baslik)
    row+=1

#fasta matrisini excele yazıyorum
row=2
col=2
for a in range(0,len(i)):
    col=2
    for b in range(0,len(j)):
        #her seferinde rasgele renk üretiyorum ve yazı formatı ayarlıyorum
        renk= '#%02X%02X%02X' % (r(),r(),r())
        eslesme = workbook.add_format({'fg_color': renk, 'color': '#FFFFFF','bold': True,'font_size':16,'align':'center'})
        #eğer matriste * bulursa yıldızlar sağ aşağıya doğru devam ettikçe aynı renge boyuyorum.
        if fasta[a][b]=="*":
            c=0
            while fasta[a+c][b+c]=="*":
                #eğer bulunan yıldızın sol üstünde yıldız varsa daha önce boyamış olduğum için break ile döngüden çıkartıyorum
                if fasta[a-1][b-1]=="*":
                    break
                worksheet.write(row+c, col+c,fasta[a+c][b+c],eslesme)
                c+=1        
        else:
            worksheet.write(row, col,fasta[a][b])
        col+=1
    row+=1
#exceli kapatıp kaydediyorum
workbook.close()