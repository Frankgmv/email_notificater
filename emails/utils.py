import win32com.client as win32
import os
import sys


def enviar_correo_analista_ro(destinatario, sr, descripcion, dias_vencido):
    try:
        outlook = win32.Dispatch("outlook.application")

        # Se crea el correo
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = f"{sr} Alerta - ¡Tienes una gestión pendiente por realizar!"

        # Se envían a estos correos como copia de seguridad
        mail.BCC = "stcolo@bancolombia.com.co;danivela@bancolombia.com.co"

        # Se crea el cuerpo del correo en HTML
        # Presionar Alt + Z para visualizar el HTML completo del cuerpo del correo
        html_body = f"""
         <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m=http://schemas.microsoft.com/office/2004/12/omml xmlns=http://www.w3.org/TR/REC-html40>
         
         <head>
         <meta http-equiv=Content-Type content="text/html; charset=utf-8">
         <meta name=Generator content="Microsoft Word 15 (filtered medium)">
         </head>
         
         <body lang=ES-CO link="#467886" vlink=purple style='word-wrap:break-word'>
         
         <div class=WordSection1><p class=MsoNormal style='margin-left:72.0pt'>&nbsp;<o:p></o:p></p><div align=center><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='max-width:450.0pt;background:white'><tr><td valign=top style='padding:0cm 0cm 0cm 0cm'><div align=center><tr></tr><tr><td width=420 style='width:315.0pt;padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'></td></tr><tr style='height:12.75pt'><td style='padding:0cm 0cm 0cm 0cm;height:12.75pt;text-size-adjust: 100%'><p class=MsoNormal style='mso-line-height-alt:.75pt'><span lang=ES-TRAD style='font-size:1.0pt'>&nbsp;</span><o:p></o:p></p></td><td style='padding:0cm 0cm 0cm 0cm;height:12.75pt'></td></tr></table></td></tr><tr><td style='padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width="100%" style='width:100.0%;text-size-adjust: 100%'><tr style='height:3.0pt'><td style='background:#FDDA24;padding:0cm 0cm 0cm 0cm;height:3.0pt;text-size-adjust: 100%'><p class=MsoNormal style='mso-line-height-alt:.75pt'><span lang=ES-TRAD style='font-size:1.0pt;color:black'>&nbsp;</span><o:p></o:p></p></td></tr></table></td></tr></table></div></td></tr></table></div><p class=MsoNormal style='margin-left:72.0pt'><span lang=ES-TRAD style='font-size:12.0pt'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><o:p></o:p></p>
         
         <div align=center><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width="90%" style='width:90.0%'><tr><td valign=top style='padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><p class=MsoNormal style='line-height:13.5pt'><span lang=ES-TRAD style='font-size:13.5pt;font-family:"Arial",sans-serif'>¡Importante!</span><o:p></o:p></p></td></tr><tr style='height:7.5pt'><td valign=top style='padding:0cm 0cm 0cm 0cm;height:7.5pt;text-size-adjust: 100%'><p class=MsoNormal style='mso-line-height-alt:.75pt'><b><span lang=ES-TRAD style='font-size:27.0pt;font-family:"Arial",sans-serif;color:black'>Alerta: SR pendiente</span></b><o:p></o:p></p></td></tr><tr style='height:31.75pt'><td style='padding:0cm 0cm 0cm 0cm;height:31.75pt'></td></tr><tr><td valign=top style='padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><p class=MsoNormal style='text-align:justify'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>¡Hola!&nbsp; Esperamos que te encuentres muy bien.</span><o:p></o:p></p><p class=MsoNormal style='text-align:justify'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>&nbsp;</span><o:p></o:p></p><p class=MsoNormal style='text-align:justify'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>Recuerda que, desde el pasado mes de mayo del <b>2024,</b> como analista de RO asignado a la iniciativa {sr} eres el <b>RESPONSABLE</b> de marcar si esta iniciativa requiere o no, la participación del área de continuidad del negocio y notificarlo, <b>dentro de los primeros 2 días después de asignado este SR.</b></span><o:p></o:p>
         
         <p class=MsoNormal style='text-align:justify'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'><b>Hemos identificado que tienes este SR pendiente y vencido en la herramienta, por favor realizar la marcación en el menor tiempo posible, de lo contrarios, estás incumpliendo con los acuerdos de servicios definidos para darle un correcto funcionamiento al proceso de gestión de riesgos en proveedores.</b></span><span style='font-size:10.0pt;font-family:"Arial",sans-serif'></span><o:p></o:p></p>
         
         <p class=MsoNormal style='text-align:justify'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'><b>Información de la iniciativa:</b></span><span style='font-size:10.0pt;font-family:"Arial",sans-serif'> </span><o:p></o:p></p><ol style='margin-top:0cm'>
         
         <li class=MsoListParagraph style='margin-left:0cm;text-align:justify;mso-list:l1 level1 lfo3'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>SR: <i>{sr}</i></span><o:p></o:p></li>
         
         <li class=MsoListParagraph style='margin-left:0cm;text-align:justify;mso-list:l1 level1 lfo3'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>Descripción: <i>{descripcion}</i></span><o:p></o:p></li>
         
         <li class=MsoListParagraph style='margin-left:0cm;text-align:justify;mso-list:l1 level1 lfo3'><span font-size:10.0pt;font-family:"Arial",sans-serif'>Días vencido: <i style='background-color:red'><b>{dias_vencido} días</b></i></span><o:p></o:p></li></ol>
         
         <p class=MsoNormal style='text-align:justify'><b><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>¿Cómo lo hago?</span></b><span style='font-size:10.0pt;font-family:"Arial",sans-serif'> </span><o:p></o:p></p><ol style='margin-top:0cm' start=1 type=1><li class=MsoListParagraph style='margin-left:0cm;text-align:justify;mso-list:l1 level1 lfo3'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>Ingresa a: <b><a href=https://nam10.safelinks.protection.outlook.com/?url=https%3A%2F%2Fapps.powerapps.com%2Fplay%2Fe%2F030e1206-2356-eb46-b90a-be6ab6655b0d%2Fa%2F324b16a7-f74a-4f28-af75-7f49ecfa6cef%3FtenantId%3Db5e244bd-c492-495b-8b10-61bfd453e423%26hint%3D112cdbec-1311-4a02-bbcb-1a90f654975f%26sourcetime%3D1718374781580%26source%3Dportal&data=05%7C02%7Cstcolo%40bancolombia.com.co%7C2704e1a96fdc40f4535208dcefc083ff%7Cb5e244bdc492495b8b1061bfd453e423%7C0%7C0%7C638648858898798273%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&sdata=BzZB4VnWNEtYrpCcsXjZ6bhbLIBwccZWybXNDiQ%2BsyY%3D&reserved=0>APB81207_Portal_Proveedores - Power Apps</a>.</b></span><o:p></o:p></li><li class=MsoListParagraph style='margin-left:0cm;text-align:justify;mso-list:l1 level1 lfo3'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>Selecciona el botón <i>"Gestión de riesgos en contratación de servicios"</i></span><o:p></o:p></li><li class=MsoListParagraph style='margin-left:0cm;text-align:justify;mso-list:l1 level1 lfo3'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>Selecciona el botón <i>"Conoce el análisis de riesgo de tu iniciativa"</i></span><o:p></o:p></li><li class=MsoListParagraph style='margin-left:0cm;text-align:justify;mso-list:l1 level1 lfo3'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>Consulta tu SR y márcalo según corresponda, como aparece en el documento adjunto.&nbsp; </span><o:p></o:p></li></ol>
         
         <p class=MsoListParagraph style='text-align:justify'><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>&nbsp;</span><o:p></o:p></p><p class=MsoNormal style='text-align:justify'><span style='font-size:7.0pt;font-family:"Arial",sans-serif'><b>Contamos con tu compromiso en la gestión de riesgos en terceros,</b></span><o:p></o:p></p>
         <p class=MsoNormal style='text-align:justify'><span style='font-size:7.0pt;font-family:"Arial",sans-serif'>Si tienes algun incidente con la herramienta puedes contactar a : Daniela Velasquez Gomez o <b>Steeven Colorado García.</b></span><o:p></o:p></p><p class=MsoNormal style='text-align:justify'><b><span lang=ES style='font-size:12.0pt;font-family:"Arial",sans-serif;color:#002060;mso-fareast-language:ES-CO'>&nbsp;</span></b><o:p></o:p></p><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Arial",sans-serif'>¡Saludos!</span></b><o:p></o:p></p></td></tr></table></div></td></tr></table></div><p class=MsoNormal style='margin-left:72.0pt'>&nbsp;<o:p></o:p></p><div align=center><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='max-width:450.0pt;background:white;text-size-adjust: 100%;font-variant-ligatures: normal;font-variant-caps: normal;orphans: 2;text-align:start;widows: 2;-webkit-text-stroke-width: 0px;text-decoration-thickness: initial;text-decoration-style: initial;text-decoration-color: initial;word-spacing:0px'><tr><td valign=top style='padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><div align=center><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='max-width:450.0pt;text-size-adjust: 100%'><tr><td valign=top style='padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><div align=center><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width="100%" style='width:100.0%;text-size-adjust: 100%'><tr><td valign=top style='padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><div align=center><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width="100%" style='width:100.0%;text-size-adjust: 100%'><tr><td width="6%" valign=top style='width:6.0%;padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><p class=MsoNormal align=center style='text-align:center'><span lang=ES-TRAD><img border=0 width=8 height=188 style='width:.0833in;height:1.9583in' id="_x0000_i1032" src=http://bancolombia-email-wsuite.s3.amazonaws.com/templates/6075b18e57ad717760ad720f/img/legal.png></span><o:p></o:p></p></td><td valign=top style='padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><div align=center><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width="100%" style='width:100.0%;text-size-adjust: 100%'><tr style='height:33.75pt'><td style='padding:7.5pt 0cm 7.5pt 0cm;height:33.75pt;text-size-adjust: 100%'><p class=MsoNormal style='mso-line-height-alt:.75pt'><span lang=ES-TRAD style='font-size:1.0pt'>&nbsp;</span><o:p></o:p></p></td></tr><tr><td valign=top style='padding:7.5pt 0cm 7.5pt 0cm;text-size-adjust: 100%'><p class=MsoNormal align=center style='text-align:center'><span lang=ES-TRAD><img border=0 width=175 height=24 style='width:1.8263in;height:.25in' id="_x0000_i1031" src=http://bancolombia-email-wsuite.s3.amazonaws.com/templates/6075b18e57ad717760ad720f/img/footer-logo.png alt=footer-logo></span><o:p></o:p></p></td></tr><tr style='height:22.5pt'><td style='padding:7.5pt 0cm 7.5pt 0cm;height:22.5pt;text-size-adjust: 100%'><p class=MsoNormal style='mso-line-height-alt:.75pt'><span lang=ES-TRAD style='font-size:1.0pt'>&nbsp;</span><o:p></o:p></p></td></tr><tr><td style='padding:0cm 0cm 7.5pt 0cm;text-size-adjust: 100%'><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width="100%" style='width:100.0%;text-size-adjust: 100%'><tr><td valign=top style='padding:0cm 0cm 0cm 0cm;text-size-adjust: 100%'><p class=MsoNormal align=center style='text-align:center'><span lang=ES-TRAD><img border=0 width=1025 height=3 style='width:10.6805in;height:.0277in' id="_x0000_i1030" src=http://bancolombia-email-wsuite.s3.amazonaws.com/templates/6075b18e57ad717760ad720f/img/line.png alt=footer-logo></span><o:p></o:p></p></td></tr></table></td></tr><tr><td style='padding:30.75pt 0cm 7.5pt 0cm;text-size-adjust: 100%'><p style='line-height:9.0pt;margin:unset'><span style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black;mso-fareast-language:EN-US'>Bancolombia nunca le solicitará datos financieros como usuarios, claves, números de tarjetas de crédito con sus códigos de seguridad y fechas de vencimiento mediante vínculos de correo electrónico o llamadas telefónicas. Para verificar la autenticidad de este correo electrónico puede reenviarlo a&nbsp;</span><a name=qLink13></a><a href=mailto:correosospechoso@bancolombia.com.co target="_blank"><span style='mso-bookmark:qLink13'><b><span style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black;mso-fareast-language:EN-US'>correosospechoso@bancolombia.com.co</span></b></span></a><span style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black;mso-fareast-language:EN-US'>.</span><o:p></o:p></p></td></tr><tr><td style='padding:16.5pt 0cm 7.5pt 0cm;text-size-adjust: 100%'><p class=MsoNormal style='line-height:9.0pt'><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>Si tiene alguna inquietud puede contactarnos en nuestras líneas de atención telefónica:<br><br>Bogotá&nbsp;</span><a name=qLink12></a><a href=tel:(571)%20343%200000><span style='mso-bookmark:qLink12'><b><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>(571) 343 0000</span></b></span></a><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>, Medellín&nbsp;</span><a name=qLink11></a><a href=tel:(574)%20510%209000><span style='mso-bookmark:qLink11'><b><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>(574) 510 9000</span></b></span></a><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>, Cali&nbsp;</span><a name=qLink10></a><a href=tel:(572)%20554%200505><span style='mso-bookmark:qLink10'><b><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>(572) 554 0505</span></b></span></a><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>, Barranquilla&nbsp;</span><a name=qLink9></a><a href=tel:(575)%20361%208888><span style='mso-bookmark:qLink9'><b><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>(575) 361 8888</span></b></span></a><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>, Bucaramanga&nbsp;</span><a name=qLink8></a><a href=tel:(577)%20697%202525><span style='mso-bookmark:qLink8'><b><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>(577) 697 2525</span></b></span></a><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>, Cartagena&nbsp;</span><a name=qLink7></a><a href=tel:(575)%20693%204400><span style='mso-bookmark:qLink7'><b><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>(575) 693 4400</span></b></span></a><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>, Resto del país&nbsp;</span><a name=qLink6></a><a href=tel:018000912345><span style='mso-bookmark:qLink6'><b><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>018000912345</span></b></span></a><span lang=ES-TRAD style='font-size:7.0pt;font-family:"Arial",sans-serif;color:black'>&nbsp;Sede principal Cra. 48 Nro. 26-85 Torre Norte. Medellín – Colombia</span><o:p></o:p></p></td></tr>
         
         </body>
         </html>
        """

        # Se establece el HTML que tendrá el cuerpo del correo "html_body"
        mail.HTMLBody = html_body

        # Se especifica la ruta del correo Outlook que se va adjuntar en el correo
        ruta_pdf_paso_a_paso = rutas + r"\Correo_Outlook.msg"

        # Se especifica la ruta del PDF que se va adjuntar en el correo
        ruta_pdf = rutas + r"\Integracion_Continuidad-RO.pdf"

        # Adjuntar el correo guardado de Outlook
        if os.path.exists(ruta_pdf_paso_a_paso):
            mail.Attachments.Add(ruta_pdf_paso_a_paso)
            print(f"\nCorreo Outlook adjunto: {ruta_pdf_paso_a_paso}")
        else:
            print(f"No se encontró el correo Outlook: {ruta_pdf_paso_a_paso}")

        # Adjuntar el PDF
        if os.path.exists(ruta_pdf):
            mail.Attachments.Add(ruta_pdf)
            print(f"\nPDF adjunto: {ruta_pdf}")
        else:
            print(f"No se encontró el PDF: {ruta_pdf}")

        # Enviar correo
        mail.Send()
        print(f"\nSe envió correctamente la información a: {destinatario}")
    except Exception as e:
        print("No se pudo enviar la información a los correos")

# TODO este es la función para aplicar mañana

def email_for_clients(client):
    correo1 = client.correo1
    correo2 = client.correo2
    nit = client.nit
    nombre_empresa = client.nombre_empresa
    nombre_gerente = client.nombre_gerente
    correo_gerente = client.correo_gerente

    my_data = {
        "nombre_remitente": "Frank Giovany Muriel Velásquez",
        "correo_remitente": "fgmuriel@bancolombia.com.co",
    }

    try:
        outlook = win32.Dispatch("outlook.application")
        destinatario = correo1 if correo2 == "" else f"{correo1}; {correo2}"

        # Se crea el correo
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.cc = correo_gerente
        mail.Subject = f"¡Tic Tac! el tiempo se acaba {nombre_empresa}, tu nueva Sucursal Virtual Negocios te espera"

        # Se crea el cuerpo del correo en HTML
        # Presionar Alt + Z para visualizar el HTML completo del cuerpo del correo
        html_body = f"""
         <h1>hola</h1>
        """
        # Se establece el HTML que tendrá el cuerpo del correo "html_body"
        mail.HTMLBody = html_body

        ruta_pdf_paso_a_paso = "aquí la ruta del paso a paso en el pc"
        # Adjuntar el correo guardado de Outlook
        # if os.path.exists(ruta_pdf_paso_a_paso):
        #     mail.Attachments.Add(ruta_pdf_paso_a_paso)
        # else:
        #     print(f"No se encontró la ruta: {ruta_pdf_paso_a_paso}")

        # Enviar correo
        # mail.Send()
        print(mail)
        print(f"Se envió correctamente la información a: {nombre_empresa} - {nit}")
        return True
    except Exception as e:
        print(e)
        # print(f"No se pudo enviar la información a la empresa {nombre_empresa} - {nit}")
