from reportlab.pdfgen import canvas 
import pandas as pd
import requests
import os


class Cartilla():
    
    def __init__(self,path):
        
        """
        La función se inicializa con el path donde está el la base de datos y retorna
        la amenaza, exposición, fragilidad, vulnerabilidad, resiliencia y riesgo para 
        cada tipo de amenaza
    
        """
        
                
        ##################### INFORMACIÓN GENERAL
        
        self.df = pd.read_excel(path,sheet_name='Escuelas')
        df0 = pd.read_excel(path,sheet_name='Evaluacion')
        
        self.inst = self.df['Institucion'].dropna().tolist()
        self.numpis = self.df['Numero de pisos'].dropna().tolist()
        self.tipolo = df0['Tipología'].dropna().tolist()
        self.ano = self.df['Año'].dropna().tolist()
        self.mun = self.df['Municipio'].dropna().tolist()
        self.tipcu = df0['Tipo de techo'].dropna().tolist()
        self.numest = self.df['Numero de Estudiantes'].tolist()
        self.subreg = df0['Subregion'].dropna().tolist()
        
        
        ######################## SISMO
        df1 = pd.read_excel(path,sheet_name = 'Sismo')

        self.insti = df1['Institucion'].dropna().tolist()

        self.ame_sis = df1['Amenaza'].tolist()
        self.expo_sis = df1['Exposición'].tolist()
        self.frag_sis = df1['Fragilidad'].tolist()
        self.vul_sis = df1['Vulnerabilidad'].tolist()
        self.Resi_sis = df1['Resiliencia'].tolist()
        self.Ries_sis = df1['Riesgo'].tolist()
        
        ###################### MOMIVIENTO EN MASA
        
        df2 = pd.read_excel(path,sheet_name = 'MMasa')
        
        self.ame_mm = df2['Amenaza'].tolist()
        self.expo_mm = df2['Exposición'].tolist()
        self.frag_mm = df2['Fragilidad'].tolist()
        self.vul_mm = df2['Vulnerabilidad'].tolist()
        self.Resi_mm = df2['Resiliencia'].tolist()
        self.Ries_mm = df2['Riesgo'].tolist()
        
        ###################### VENDAVAL
        
        df3 = pd.read_excel(path,sheet_name = 'Vendaval')
        
        self.ame_ven = df3['Amenaza'].tolist()
        self.expo_ven = df3['Exposición'].tolist()
        self.frag_ven = df3['Fragilidad'].tolist()
        self.vul_ven = df3['Vulnerabilidad'].tolist()
        self.Resi_ven = df3['Resiliencia'].tolist()
        self.Ries_ven = df3['Riesgo'].tolist()
        
        ###################### AVENIDAS TORRENCIALES
        
        df4 = pd.read_excel(path,sheet_name = 'ATorrenciales')
        
        self.ame_at = df4['Amenaza'].tolist()
        self.expo_at = df4['Exposición'].tolist()
        self.frag_at = df4['Fragilidad'].tolist()
        self.vul_at = df4['Vulnerabilidad'].tolist()
        self.Resi_at = df4['Resiliencia'].tolist()
        self.Ries_at = df4['Riesgo'].tolist()
        
        ###################### RIESGO MULTIAMENAZA
        
        df5 = pd.read_excel(path,sheet_name = 'RM')

        self.r_m = df5['Riesgo multiamenaza'].tolist()

        ###################### PARÁMETROS QUE INFLUYEN EL RIESGO

        df6 = pd.read_excel(path,sheet_name = 'F_Sis')
        self.peso_sis =df6['Peso'].tolist()
        self.param_sis = df6['Parámetro'].tolist()


        df7 = pd.read_excel(path,sheet_name = 'F_MMasa')
        self.peso_mm =df7['Peso'].tolist()
        self.param_mm = df7['Parámetro'].tolist()

        df8 = pd.read_excel(path,sheet_name = 'F_Ven')
        self.peso_ven =df8['Peso'].tolist()
        self.param_ven = df8['Parámetro'].tolist()

        df9 = pd.read_excel(path,sheet_name = 'F_Ator')
        self.peso_at =df9['Peso'].tolist()
        self.param_at = df9['Parámetro'].tolist()

        return 


    def Ima(self):

        img = {"índice Normalizado":self.df['Índice Normalizado'].dropna(),"Columna11":self.df['Columna11'].dropna(),"Fachada Derecha":self.df['Fachada Derecha Img'].dropna(),"Fachada Posterior":self.df['Fachada Posterior Img'].dropna(),"Foto Escuela":self.df['Foto Escuela'].dropna()}
        images = pd.DataFrame(img)

        link = []
        for i1 in range(len(images)):

            l=images.iloc[i1].dropna().tolist()
            link.append(l)

        for i2 in range(len(self.inst)):

            nombre_carpeta = self.inst[i2].strip()
            carpeta = os.path.join("C:\\Users\\Usuario\\Desktop\\Aleja\\Fotos_es",nombre_carpeta )
                
            if not os.path.exists(carpeta):    
                    os.makedirs(carpeta)  # crea la cantidad de carpetas para las fotos
            for j in range(len(link[i2])):
                    
                
                a = requests.get(link[i2][j])
                with open(os.path.join(carpeta,f"imagen_{j}.jpg"), "wb") as f:
                        f.write(a.content)

        return 

    def plan(self):
        imgs = []
        for i in range(len(self.insti)):
            width1 , height1 = (215* 2.83464567,330* 2.83464567)
            c = canvas.Canvas ('C:/Users/Usuario/Desktop/Aleja/Cartillas/{}.pdf'.format(self.insti[i]),pagesize = (width1,height1))
            
            mc = 255

        ################### ERM TEXT and IMAGE ERM

            

            path_um = 'C:/Users/Usuario/Desktop/todas las carpetas/Clases 2021-1/Fotos Látex/Universidad-de-Medellín-UDEM-logo.png'      
            path_un = 'C:/Users/Usuario/Desktop/todas las carpetas/Clases 2021-1/Fotos Látex/logo_un.PNG'
            path_eia = 'C:/Users/Usuario/Desktop/todas las carpetas/Clases 2021-1/Fotos Látex/Logo_EIA.PNG'


            c.drawImage('ERM.PNG',x=20,y=820,width = 100, height = 100)
            c.drawImage(path_um,x=550,y=10,width = 40, height = 40)
            c.drawImage(path_un,x=480,y=10,width = 60, height = 50)
            c.drawImage(path_eia,x=400,y=10,width = 60, height = 40)



            c.setFont('Helvetica-Bold', size = 50)
            c.setFillColorRGB(91/mc,119/mc,128/mc)
            c.drawString(130,845,'ERM')
            
            
        ######################## elipses info
            
            
            ancho_insti = c.stringWidth(str(self.insti[i]), fontName='Helvetica-Bold', fontSize=8)
            ancho_np = c.stringWidth(str(self.numpis[i]), fontName='Helvetica-Bold', fontSize=8)
            ancho_tipo = c.stringWidth(str(self.tipolo[i]), fontName='Helvetica-Bold', fontSize=8)
            ancho_mun = c.stringWidth(str(self.mun[i]), fontName='Helvetica-Bold', fontSize=8)
            ancho_tc = c.stringWidth(str(self.tipcu[i]), fontName='Helvetica-Bold', fontSize=8)
            ancho_ano = c.stringWidth(str(self.ano[i]), fontName='Helvetica-Bold', fontSize=8)
            ancho_ne = c.stringWidth(str(int(self.numest[i])), fontName='Helvetica-Bold', fontSize=8)
            ancho_sr = c.stringWidth(str(self.subreg[i]), fontName='Helvetica-Bold', fontSize=8)
            

            c.setStrokeColorRGB(91/mc,119/mc,128/mc)
            c.setFillColorRGB(91/mc,119/mc,128/mc)

            c.roundRect(x=30,y=770,width =50, height = 15,radius=5,fill=1)
            c.roundRect(x=30,y=750,width =80, height = 15,radius=5,fill=1)
            c.roundRect(x=30,y=730,width =70, height = 15,radius=5,fill=1)


            c.roundRect(x=ancho_insti+100,y=770,width =50, height = 15,radius=5,fill=1)
            c.roundRect(x=ancho_np+180,y=750,width =80, height = 15,radius=5,fill=1)
            c.roundRect(x=ancho_tipo+120,y=730,width =90, height = 15,radius=5,fill=1)

            c.roundRect(x=ancho_insti+170+ancho_ano,y=770,width =110, height = 15,radius=5,fill=1)
            c.roundRect(x=ancho_np+280+ancho_mun,y=750,width =70, height = 15,radius=5,fill=1)

            c.setStrokeColorRGB(mc,mc,mc)
            c.setFillColorRGB(mc,mc,mc)
            c.setFont('Helvetica-Bold', size = 12)

            c.drawString(40,775,'CER')
            c.drawString(40,755,'N° de pisos')
            c.drawString(40,735,'Tipología')
            
            c.drawString(ancho_insti+110,775,'Año')
            c.drawString(ancho_np+190,755,'Municipio')
            c.drawString(ancho_tipo+130,735,'Tipo cubierta')

            c.drawString(ancho_insti+175+ancho_ano,775,'N° de estudiantes')
            c.drawString(ancho_np+285+ancho_mun,755,'Subregión')

            
            
            # En este apartado va el texto de la información General que varía 
            
            
            
            
            
            ###### CER, NUMERO DE PISOS, TIPOLOGÍA RECTÁNGULOS

            
            c.saveState()
            c.setStrokeColorRGB(mc, mc, mc)
            c.setFillColorRGB(132/mc, 172/mc, 184/mc,alpha = 0.3)
            
            c.roundRect(x=90,y=770,width =ancho_insti, height = 15,radius=5,fill=1)
            c.roundRect(x=120,y=750,width =ancho_np+30, height = 15,radius=5,fill=1)
            c.roundRect(x=110,y=730,width =ancho_tipo, height = 15,radius=5,fill=1)
            

            ###### AÑO, MUNICIPIO, TIPO DE CUBIERTA RECTÁNGULOS
            
            c.roundRect(x=ancho_insti+160,y=770,width =ancho_ano, height = 15,radius=5,fill=1)
            c.roundRect(x=ancho_np+270,y=750,width =ancho_mun, height = 15,radius=5,fill=1)
            c.roundRect(x=ancho_tipo+220,y=730,width =ancho_tc, height = 15,radius=5,fill=1)

            ###### SUBREGIÓN Y NÚMERO DE ESTUDIANTES RECTÁNGULOS

            c.roundRect(x=ancho_insti+290+ancho_ano,y=770,width =ancho_ne+30, height = 15,radius=5,fill=1)
            c.roundRect(x=ancho_np+360+ancho_mun,y=750,width =ancho_sr, height = 15,radius=5,fill=1)

            c.restoreState()
            
            ###### CER, NUMERO DE PISOS, TIPOLOGÍA TEXTO
            
            c.setStrokeColorRGB(91/mc,119/mc,128/mc)
            c.setFillColorRGB(91/mc,119/mc,128/mc)
            c.setFont('Helvetica-Bold', size = 8)

            c.drawString(90,775,str(self.inst[i]))
            c.drawString(135,755,str(self.numpis[i]))
            c.drawString(110,735,str(self.tipolo[i]))
            
            ###### AÑO, MUNICIPIO, TIPO DE CUBIERTA TEXTO
            
            c.drawString(ancho_insti+160,775,self.ano[i])
            c.drawString(ancho_np+270,755,self.mun[i])
            c.drawString(ancho_tipo+220,735,self.tipcu[i])
            
            ###### SUBREGIÓN Y NÚMERO DE ESTUDIANTES TEXTO
            
            c.drawString(ancho_insti+305+ancho_ano,775,str(int(self.numest[i])))
            c.drawString(ancho_np+360+ancho_mun,755,self.subreg[i])

        ###################### Información general text

            c.saveState()
            c.setStrokeColorRGB(mc, mc, mc)
            c.setFillColorRGB(91/mc,119/mc,128/mc)
            c.rect(20,800,570,20,fill = 1)
            c.rect(20,(height1/2)-50,570,20,fill = 1)

            c.setFont('Helvetica-Bold', size = 12)
            c.setStrokeColorRGB(mc, mc, mc)
            c.setFillColorRGB(mc,mc,mc)
            c.drawCentredString(width1/2,805,'Información General')
            c.drawCentredString(width1/2,(height1/2)-45,'Evaluación del Riesgo')

            c.restoreState()
            c.saveState()
            c.setFillColorRGB(132/mc, 172/mc, 184/mc,alpha = 0.3)
            c.setStrokeColorRGB(91/mc,119/mc,128/mc)
            
            c.rect(20,380,60,20,fill=1)
            c.rect(80,380,60,20,fill=1)
            c.rect(140,380,60,20,fill=1)
            c.rect(200,380,60,20,fill=1)
            c.rect(260,380,60,20,fill=1)
            c.rect(320,380,60,20,fill=1)
            c.rect(380,380,60,20,fill=1)

            c.rect(20,360,60,20,fill=1)
            c.rect(20,340,60,20,fill=1)
            c.rect(20,320,60,20,fill=1)
            c.rect(20,300,60,20,fill=1)

            c.rect(460,340,130,60,fill=1)
            
            c.restoreState()


            
    
            
            c.saveState()
            c.setFont('Helvetica-Bold', size = 8)
            c.setStrokeColorRGB(91/mc,119/mc,128/mc)
            c.setFillColorRGB(91/mc,119/mc,128/mc)
            c.drawString(40,387,'Tipo')
            c.drawString(92,387,'Amenaza')
            c.drawString(148,387,'Exposición')
            c.drawString(210,387,'Fragilidad')
            c.drawString(262,387,'Vulnerabilidad')
            c.drawString(330,387,'Resiliencia')
            c.drawString(396,387,'Riesgo')
            
            c.drawString(38,367,'Sismo')
            c.drawString(24,347,'Mov. en masa')
            c.drawString(22,327,'Ave. torrencial')
            c.drawString(32,307,'Vendaval')
            
            c.setFont('Helvetica-Bold', size = 10)
            c.drawString(510,374,'Riesgo')
            c.drawString(495,364,'Multiamenaza')
            
            c.restoreState()
            
    ###################### Ubiación y registro fotográfico text

            c.saveState()

            c.setStrokeColorRGB(mc, mc, mc)
            c.setFillColorRGB(255/mc, 133/mc, 10/mc)
            c.rect(20,700,200,20,fill = 1)
            c.rect(225,700,364,20,fill =1)
            c.rect(20,(height1/2)-200,364,20,fill = 1)
            c.rect(389,(height1/2)-200,200,20,fill = 1)
            
            
            c.setFont('Helvetica-Bold', size = 12)
            c.setFillColorRGB(mc,mc,mc)
            c.setStrokeColorRGB(mc,mc,mc)
            c.drawString(100,705,'Ubación')
            c.drawString(350,705,'Registro fotográfico')
            c.drawString(110,(height1/2)-195,'Parámetros que incrementan el riesgo')
            c.drawString(435,(height1/2)-195,'Recomendaciones')
            
            
    ####---------------------------------------------

            c.setStrokeColorRGB(255/mc, 133/mc, 10/mc)
            c.rect(20,(height1/2)-25,198,250)
            
                    # ESPACIO MAPA
                    
            c.rect(20,20,362,240)

                # ESPACIO RIESGOS
                
            c.rect(389,60,200,200)

                # ESPACIO RECOMENDACIONES
                
            c.restoreState()
    ####---------------------------------------------




    ####################### Evaluación riesgo text
            
            c.saveState()
            
            c.setFont('Helvetica-Bold', size = 16)
            c.setFillColorRGB(91/mc,119/mc,128/mc)
            c.setStrokeColorRGB(91/mc,119/mc,128/mc)
            c.drawString(275,860,'Evaluación del Riesgo Multiamenaza')
            
            c.setStrokeColorRGB(mc, mc, mc)
            c.setFillColorRGB(132/mc, 172/mc, 184/mc,alpha = 0.3)
            c.rect(251,855,339,20,fill = 1)
            
            c.restoreState()
            

############################# DIBUJAR PARÁMETROS QUE INCREMENTAN EL RIESGO 
            combinaciones = []
            for f1 in range(len(self.ame_sis)):

                com = [self.ame_sis[f1],self.ame_mm[f1],self.ame_ven[f1],self.ame_at[f1]]
                combinaciones.append(com)

            contCombi = []
            for combi in combinaciones:
                
                contAlt=combi.count("Alta")
                contInt=combi.count("Intermedia")
                contCombi.append(contAlt+contInt)

            for contCon in contCombi:

                if contCon == 1:
                    c.rect(20,20,362,20,fill = 1) ### este rectángulo se van a llenar dependiendo de los parámetros que aumenten el riesgo
                    
                
                elif contCon == 2:

                    c.rect(20,20,362,20,fill = 1) ### este rectángulo se van a llenar dependiendo de los parámetros que aumenten el riesgo
                    c.rect(20,140,362,20,fill = 1) ### este rectángulo se van a llenar dependiendo de los parámetros que aumenten el riesgo
            
                
                elif contCon == 3:

                    c.rect(20,20,362,20,fill = 1) ### este rectángulo se van a llenar dependiendo de los parámetros que aumenten el riesgo
                    c.rect(20,140,362,20,fill = 1) ### este rectángulo se van a llenar dependiendo de los parámetros que aumenten el riesgo
                    c.rect(20,140,362,20,fill = 1) ### este rectángulo se van a llenar dependiendo de los parámetros que aumenten el riesgo
            

                elif contCon == 4:

                    pass

            
############################# DIBUJAR LAS FOTOS EN LA CARTILLA 
            #########################################################

            
            a = os.listdir(f"C:\\Users\\Usuario\\Desktop\\Aleja\\Fotos_es\\{self.inst[i].strip()}") 
                
            img = []
            for j1 in range(len(a)):

                img.append(os.path.join(f"C:\\Users\\Usuario\\Desktop\\Aleja\\Fotos_es\\{self.inst[i].strip()}",a[j1]))
            
            imgs.append(img)


            
            fot2_1 = (228,(height1/2)+105,359,120)
            fot2_2 = (228,(height1/2)-25,359,120)
            fot3_1 = (228,(height1/2)+105,359,120)
            fot3_2 = (228,(height1/2)-25,174.5,120)
            fot3_3 = (412.5,(height1/2)-25,174.5,120)
            fot4_1 = (228,(height1/2)+105,174.5,120)
            fot4_2 = (415.5,(height1/2)+105,174.5,120)
            fot4_3 = (228,(height1/2)-25,174.5,120)
            fot4_4 = (415.2,(height1/2)-25,174.5,120)
            fot5_1 = (228,(height1/2)+105,235.4,120)
            fot5_2 = (471.4,(height1/2)+105,113.7,120)
            fot5_3 = (228,(height1/2)-25,113.7,120)
            fot5_4 = (349.7,(height1/2)-25,113.7,120)
            fot5_5 = (471.4,(height1/2)-25,113.7,120)

            fot1 = [(228,(height1/2)-25,359,250)]
            fot2 = [fot2_1,fot2_2]
            fot3 = [fot3_1,fot3_2,fot3_3]
            fot4 = [fot4_1,fot4_2,fot4_3,fot4_4]
            fot5 = [fot5_1,fot5_2,fot5_3,fot5_4,fot5_5]
            
            
            #x,y,width,height = res_rec_at
            for j2 in imgs: ####### esta parte crea el diccionario para enlazar las imágenes para la escuelita correcta
                
                if len (j2) < 2:

                    f1 = dict(zip(j2,fot1))
                    

                elif len (j2) < 3:

                    f2 = dict(zip(j2,fot2))
            
                elif len (j2) < 4:
                    
                    f3 = dict(zip(j2,fot3))

                elif len (j2) < 5:

                    f4 = dict(zip(j2,fot4))
            
                elif len (j2) < 6:

                    f5 = dict(zip(j2,fot5))
            
            for j3 in range(len(img)):

                if len(img) < 2:
                    x1,y1,w,h = f1[img[j3]]
                    c.saveState()
                    c.drawImage(img[j3],x=x1,y=y1,width = w, height =h)
                    c.restoreState()
                elif len(img) < 3 and len(img)> 1:
                    x1,y1,w,h = f2[img[j3]]

                    c.drawImage(img[j3],x=x1,y=y1,width = w, height =h)

                elif len(img) < 4 and len(img)> 2:
                    x1,y1,w,h = f3[img[j3]]

                    c.drawImage(img[j3],x=x1,y=y1,width = w, height =h)

                elif len(img) < 5 and len(img)> 3:
                    x1,y1,w,h = f4[img[j3]]

                    c.drawImage(img[j3],x=x1,y=y1,width = w, height =h)

            
                elif len(img) < 6 and len(img)> 4:
                    x1,y1,w,h = f5[img[j3]]

                    c.drawImage(img[j3],x=x1,y=y1,width = w, height =h)

            ################################# acá comienza la evaluación del riesgo
            s1 = (80,360,60,20)
            s2 = (140,360,60,20)
            s3 = (200,360,60,20)
            s4 = (260,360,60,20)
            s5 = (320,360,60,20)
            s6 = (380,360,60,20)
            
            mm1 = (80,340,60,20)
            mm2 = (140,340,60,20)
            mm3 = (200,340,60,20)
            mm4 = (260,340,60,20)
            mm5 = (320,340,60,20)
            mm6 = (380,340,60,20)
            
            at1 = (80,320,60,20)
            at2 = (140,320,60,20)
            at3 = (200,320,60,20)
            at4 = (260,320,60,20)
            at5 = (320,320,60,20)
            at6 = (380,320,60,20)
            
            ven1 = (80,300,60,20)
            ven2 = (140,300,60,20)
            ven3 = (200,300,60,20)
            ven4 = (260,300,60,20)
            ven5 = (320,300,60,20)
            ven6 = (380,300,60,20)
            
            rec_sis =[s1,s2,s3,s4,s6]
            rec_mm = [mm1,mm2,mm3,mm4,mm6]
            rec_at = [at1,at2,at3,at4,at6]
            rec_ven = [ven1,ven2,ven3,ven4,ven6]
            
            sis = [self.ame_sis,self.expo_sis,self.frag_sis,self.vul_sis,self.Ries_sis]
            das = dict(zip(rec_sis,sis))
            
            mm = [self.ame_mm,self.expo_mm,self.frag_mm,self.vul_mm,self.Ries_mm]       
            damm = dict(zip(rec_mm,mm))

            at = [self.ame_at,self.expo_at,self.frag_at,self.vul_at,self.Ries_at]
            dat = dict(zip(rec_at,at))

            ven = [self.ame_ven,self.expo_ven,self.frag_ven,self.vul_ven,self.Ries_ven] 
            daven = dict(zip(rec_ven,ven))

            #### Este apartado llena solo resiliencia:

            res_rec_sis= s5
            res_rec_mm= mm5
            res_rec_at= at5
            res_rec_ven= ven5

            resi_sis = self.Resi_sis
            resi_mm = self.Resi_mm
            resi_at = self.Resi_at
            resi_ven = self.Resi_ven

            if resi_sis[i]=='Baja':
                x,y,width,height = res_rec_sis
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+20,y+5,'Baja')
                c.restoreState()
            elif resi_sis[i] == 'Intermedia':
                x,y,width,height = res_rec_sis
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+10,y+5,'Intermedia')
                c.restoreState()
            elif resi_sis[i]=='Alta':
                x,y,width,height = res_rec_sis
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+20,y+5,'Alta')
                c.restoreState()

            if resi_mm[i]=='Baja':
                x,y,width,height = res_rec_mm
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+20,y+5,'Baja')
                c.restoreState()
            elif resi_mm[i] == 'Intermedia':
                x,y,width,height = res_rec_mm
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+10,y+5,'Intermedia')
                c.restoreState()
            elif resi_mm[i]=='Alta':
                x,y,width,height = res_rec_mm
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+20,y+5,'Alta')
                c.restoreState()

            if resi_at[i]=='Baja':
                x,y,width,height = res_rec_at
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+20,y+5,'Baja')
                c.restoreState()
            elif resi_at[i] == 'Intermedia':
                x,y,width,height = res_rec_at
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+10,y+5,'Intermedia')
                c.restoreState()
            elif resi_at[i]=='Alta':
                x,y,width,height = res_rec_at
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+20,y+5,'Alta')
                c.restoreState()

            if resi_ven[i]=='Baja':
                x,y,width,height = res_rec_ven
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+20,y+5,'Baja')
                c.restoreState()
            elif resi_ven[i] == 'Intermedia':
                x,y,width,height = res_rec_ven
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+10,y+5,'Intermedia')
                c.restoreState()
            elif resi_ven[i]=='Alta':
                x,y,width,height = res_rec_ven
                c.saveState()
                c.setFont('Helvetica-Bold', size = 8)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(x+20,y+5,'Alta')
                c.restoreState()
            ####

            for x1 in rec_sis:
                
                if das[x1][i] == 'Baja':
                    
                    x,y,width,height = x1
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+20,y+5,'Baja')
                    c.restoreState()
                elif das[x1][i] == 'Intermedia':
                    x,y,width,height = x1
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+10,y+5,'Intermedia')
                    c.restoreState()
                elif das[x1][i] == 'Alta':
                    x,y,width,height = x1
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+20,y+5,'Alta')
                    c.restoreState()


            for x2 in rec_mm:
                
                if damm[x2][i] == 'Baja':
                    
                    x,y,width,height = x2
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+20,y+5,'Baja')
                    c.restoreState()
                elif damm[x2][i] == 'Intermedia':
                    x,y,width,height = x2
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+10,y+5,'Intermedia')
                    c.restoreState()
                elif damm[x2][i] == 'Alta':
                    x,y,width,height = x2
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+20,y+5,'Alta')
                    c.restoreState()

            for x3 in rec_at:
                
                if dat[x3][i] == 'Baja':
                    
                    x,y,width,height = x3
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+20,y+5,'Baja')
                    c.restoreState()
                elif dat[x3][i] == 'Intermedia':
                    x,y,width,height = x3
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+10,y+5,'Intermedia')
                    c.restoreState()
                elif dat[x3][i] == 'Alta':
                    x,y,width,height = x3
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+20,y+5,'Alta')
                    c.restoreState()

            for x4 in rec_ven:
                
                if daven[x4][i] == 'Baja':
                    
                    x,y,width,height = x4
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+20,y+5,'Baja')
                    c.restoreState()
                elif daven[x4][i] == 'Intermedia':
                    x,y,width,height = x4
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+10,y+5,'Intermedia')
                    c.restoreState()
                elif daven[x4][i] == 'Alta':
                    x,y,width,height = x4
                    c.saveState()
                    c.setFont('Helvetica-Bold', size = 8)
                    c.setFillColorRGB(91/mc,119/mc,128/mc)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.drawString(x+20,y+5,'Alta')
                    c.restoreState()


            if self.r_m[i] == 'Baja':
                c.setFont('Helvetica-Bold', size = 10)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(515,315,'Baja')
            elif self.r_m[i] == 'Intermedia':
                c.setFont('Helvetica-Bold', size = 10)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(500,315,'Intermedia')
            elif self.r_m[i]== 'Alta':
                c.setFont('Helvetica-Bold', size = 10)
                c.setFillColorRGB(91/mc,119/mc,128/mc)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.drawString(515,315,'Alta')

            for i1 in rec_sis:
                
                if das[i1][i] == 'Baja':
                    
                    x,y,width,height = i1
                    c.saveState()
                    c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    c.rect(x,y,width,height,fill=1)
                    c.restoreState()
                    
                    
                elif das[i1][i] == 'Intermedia':
                    c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    x,y,width,height = i1
                    c.rect(x,y,width,height,fill=1)
                    
                    
                elif das[i1][i] == 'Alta':
                    c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    x,y,width,height = i1
                    c.rect(x,y,width,height,fill=1)
                
                    
            for i2 in rec_mm:
                        
                if damm[i2][i] == 'Baja':
                    c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    
                    x,y,width,height = i2
                    c.rect(x,y,width,height,fill=1)
                    
                    

                    
                elif damm[i2][i] == 'Intermedia':
                    c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    x,y,width,height = i2
                    c.rect(x,y,width,height,fill=1)
                elif damm[i2][i] == 'Alta':
                    c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    x,y,width,height = i2
                    c.rect(x,y,width,height,fill=1)

            for i3 in rec_at:
                        
                if dat[i3][i] == 'Baja':
                    c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                            
                    x,y,width,height = i3
                    c.rect(x,y,width,height,fill=1)
                    
                elif dat[i3][i] == 'Intermedia':
                    c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    x,y,width,height = i3
                    c.rect(x,y,width,height,fill=1)
                elif dat[i3][i] == 'Alta':
                    c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    x,y,width,height = i3
                    c.rect(x,y,width,height,fill=1)
                    


            for i4 in rec_ven:
                        
                if daven[i4][i] == 'Baja':
                    c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                            
                    x,y,width,height = i4
                    c.rect(x,y,width,height,fill=1)
                    
                elif daven[i4][i] == 'Intermedia':
                    c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    x,y,width,height = i4
                    c.rect(x,y,width,height,fill=1)
                elif daven[i4][i] == 'Alta':
                    c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                    c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                    x,y,width,height = i4
                    c.rect(x,y,width,height,fill=1)
            

            
            if self.r_m[i] == 'Baja':
                c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(460,300,130,40,fill = 1)
            elif self.r_m[i] == 'Intermedia':
                c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(460,300,130,40,fill = 1)
            elif self.r_m[i]== 'Alta':
                c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(460,300,130,40,fill = 1)
            # En este apartado entra los cuadros de información del riesgo que cambia RECTANGULOS

           #### Este apartado llena solo resiliencia:

            if resi_sis[i]=='Baja':
                x,y,width,height = res_rec_sis
                c.saveState()
                c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()
            elif resi_sis[i] == 'Intermedia':
                x,y,width,height = res_rec_sis
                c.saveState()
                c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()
            elif resi_sis[i]=='Alta':
                x,y,width,height = res_rec_sis
                c.saveState()
                c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()

            if resi_mm[i]=='Baja':
                x,y,width,height = res_rec_mm
                c.saveState()
                c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()
            elif resi_mm[i] == 'Intermedia':
                x,y,width,height = res_rec_mm
                c.saveState()
                c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()
            elif resi_mm[i]=='Alta':
                x,y,width,height = res_rec_mm
                c.saveState()
                c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()

            if resi_at[i]=='Baja':
                x,y,width,height = res_rec_at
                c.saveState()
                c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()
            elif resi_at[i] == 'Intermedia':
                x,y,width,height = res_rec_at
                c.saveState()
                c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()
            elif resi_at[i]=='Alta':
                x,y,width,height = res_rec_at
                c.saveState()
                c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()

            if resi_ven[i]=='Baja':
                x,y,width,height = res_rec_ven
                c.saveState()
                c.setFillColorRGB(245/mc, 93/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()
            elif resi_ven[i] == 'Intermedia':
                x,y,width,height = res_rec_ven
                c.saveState()
                c.setFillColorRGB(245/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()
            elif resi_ven[i]=='Alta':
                x,y,width,height = res_rec_ven
                c.saveState()
                c.setFillColorRGB(90/mc, 245/mc, 66/mc,alpha = 0.3)
                c.setStrokeColorRGB(91/mc,119/mc,128/mc)
                c.rect(x,y,width,height,fill = 1)
                c.restoreState()

            

            
            

            c.save()
            
        return print(contCombi)
            
            
c1 = Cartilla('C:/Users/Usuario/Desktop/Aleja/Base de Datos.xlsx')
c1.plan()
#c1.crear_Carpetas()
#c1.Ima()