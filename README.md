<div align="center">

## CMHTML

<img src="globe.jpg">
</div>

### Description

This class demonstrates how to send MHTML email. It has a very simple interface which allows you to send HTML formatted email with 2 method calls. This combines code from Sebastian, Luis Cantero (EncodeBase64) and Brian Anderson for the SMTP Winsock code. Sample code is included.
 
### More Info
 
There are 2 method calls required:

1) AddAttachment - which takes the full pathname of file you want to attach and also the content identifier (CID) which you will use in your HTML email body

2) SendEmail - which takes 6 parameters (see example).

MHTML is used by most of the newish email clients including Outlook Express and many of the web email sites. It allows you to include all the richness of HTML into your email message which is very cool. For more information regarding MHTML please lookup the following site: http://www.dsv.su.se/~jpalme/ietf/web-email.html

AddAttachment returns true if file to attach is accepted and False if it does not exist

SendEmail returns true if email is sent ok and false otherwise.


<span>             |<span>
---                |---
**Submitted On**   |2000-01-24 23:42:06
**By**             |[Gary Ong](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gary-ong.md)
**Level**          |Advanced
**User Rating**    |4.7 (47 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD30291242000\.zip](https://github.com/Planet-Source-Code/gary-ong-cmhtml__1-5658/archive/master.zip)








