Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1
Dim strErrLog : strErrLog = ""
Dim bFlagDownloadAttachmentsWithFooter: bFlagDownloadAttachmentsWithFooter = True ' Set this variable to False to download images without the footer
Dim bFlagDownloadAttachmentsWithOriginalName : bFlagDownloadAttachmentsWithOriginalName = False  'Set this variable to True to download attachments with the original filename
Dim bFlagDownloadAttachmentsWithSpecialName : bFlagDownloadAttachmentsWithSpecialName = False  'Set this variable to True to download attachments with the special filename format - ["InstanceID_Location City_Store ID_Visit Date_Attachment Position ObjectName"]
Dim downloadInFolders: downloadInFolders = False 'Set this to True to download attachments in folders grouped by LocationID_Address_City
Dim foldersNameTemplate : foldersNameTemplate = "" 'You can add your folder template name in this string. You can add custom text and columns from the database. The columns must be surrounded with the symbol #. You can choose from the following columns:
'			#LocationName#
'			#LocationID#
'			#City#
'			#ClientID#
'			#LocationCountryCode#
'			#Address#
'Examples:
'	Dim attNameTemplate : attNameTemplate = "#LocationID#_#LocationName#_#ClientID#"
Dim attNameTemplate : attNameTemplate = "#ClientID#_#SurveyInstanceID#_#LocationCountryCode#_#AttachmentPosition#_#AttachmentID#_#LoopObjectName#_#ObjectName#_#FileName#" 'You can add your attachment template name in this string. You can add custom text and columns from the database. The columns must be surrounded with the symbol #. You can choose from the following columns:
'			#LocationName#
'			#AttachmentID#
'			#SurveyInstanceID# 
'			#LocationID# 
'			#AttachmentPosition#
'			#FormID#
'			#City#
'			#ObjectName#
'			#LoopObjectName#
'			#FileName#
'			#QuestionID#
'			#VisitDate#
'			#QuestionText#
'			#SectionTitle#
'			#SubsectionTitle#
'			#AttachmentSequence#
'			#ClientID#
'			#LocationCountryCode#
'Examples:
'	Dim attNameTemplate : attNameTemplate = "#LocationID##LocationName##AttachmentPosition#"
'	Dim attNameTemplate : attNameTemplate = "#FormID#_#SurveyInstanceID#_#AttachmentPosition#"
'	Dim attNameTemplate : attNameTemplate = "#ObjectName#__CustomText_#LocationID#_#AttachmentID#"


Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033592&ImageType=original&password=D9B55BE6A9FF405A9B6F0DF4C6CE89C9A118D2E3B79F4E9B81E2F2917A434D34", UNESCAPE("HN001_2024-10-09_%5B26_1766894_5826%5D.JPG"), "IMG20241009_181832.jpg", "1766894_HN_HN001_2024-10-09_26__IMAGE_PRODUCT.JPG", "jpg", "11033592", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030885&ImageType=original&password=F304A1759D1943009B292396E948E6A9406173E565D449A3BFBA3F73AA3B57E6", UNESCAPE("HN001_2024-10-09_%5B31_1766681_5826%5D.JPG"), "IMG20241009_171749.jpg", "1766681_HN_HN001_2024-10-09_31__IMAGE_PRODUCT.JPG", "jpg", "11030885", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033084&ImageType=original&password=526A11A9A67D4DCDB196EE421B1D379B97F5D9DF96A54884BB85C40CE8F55E3E", UNESCAPE("HN004_2024-10-09_%5B26_1766828_5826%5D.JPG"), "IMG20241009_164943.jpg", UNESCAPE("1766828_B%u1EAEC%20GIANG_HN004_2024-10-09_26__IMAGE_PRODUCT.JPG"), "jpg", "11033084", UNESCAPE("HN004_UBND%20TP.Bac%20Giang_B%u1EAEC%20GIANG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11032679&ImageType=original&password=2E31BD77974F44CAAB2EBC614C74BA0D051D4404C9884C1783EBDFEE513D1EE8", UNESCAPE("HN005_2024-10-09_%5B26_1766804_5826%5D.JPG"), "IMG20241009_164753.jpg", UNESCAPE("1766804_H%u1EA2I%20D%u01AF%u01A0NG_HN005_2024-10-09_26__IMAGE_PRODUCT.JPG"), "jpg", "11032679", UNESCAPE("HN005_UBND%20TP.Hai%20Duong_H%u1EA2I%20D%u01AF%u01A0NG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11032845&ImageType=original&password=755B374F9B8C412B833CD92C5EE1E4D52A6CE44140DE4345BBED81F7693392CF", UNESCAPE("HN005_2024-10-09_%5B33_1766815_5826%5D.JPG"), "IMG20241009_162417.jpg", UNESCAPE("1766815_H%u1EA2I%20D%u01AF%u01A0NG_HN005_2024-10-09_33__IMAGE_PRODUCT.JPG"), "jpg", "11032845", UNESCAPE("HN005_UBND%20TP.Hai%20Duong_H%u1EA2I%20D%u01AF%u01A0NG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030350&ImageType=original&password=286B199263674938B5840BA0DC2134F8E9E968D83BFA49A6B02FE7561040C7B6", UNESCAPE("HN007_2024-10-09_%5B28_1766570_5826%5D.JPG"), "IMG20241009_165101.jpg", UNESCAPE("1766570_THANH%20H%D3A_HN007_2024-10-09_28__IMAGE_PRODUCT.JPG"), "jpg", "11030350", UNESCAPE("HN007_UBND%20Tp.Thanh%20H%F3a_THANH%20H%D3A"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030738&ImageType=original&password=E8A8583AC07245E1BA8F66209F71E65510DACDCA4BA643EBA37166B17D3489A8", UNESCAPE("HN009_2024-10-09_%5B27_1766593_5826%5D.JPG"), "IMG20241009_165543.jpg", UNESCAPE("1766593_NGH%u1EC6%20AN_HN009_2024-10-09_27__IMAGE_PRODUCT.JPG"), "jpg", "11030738", UNESCAPE("HN009_UBND%20TP.Vinh_NGH%u1EC6%20AN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11032595&ImageType=original&password=C685944A41364E0E90FECB61D338B2E9940498A418DC4CA4A3C0359888F142A6", UNESCAPE("HN005_2024-10-09_%5B26_1766801_5826%5D.JPG"), "IMG20241009_163953.jpg", UNESCAPE("1766801_H%u1EA2I%20D%u01AF%u01A0NG_HN005_2024-10-09_26__IMAGE_PRODUCT.JPG"), "jpg", "11032595", UNESCAPE("HN005_UBND%20TP.Hai%20Duong_H%u1EA2I%20D%u01AF%u01A0NG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033622&ImageType=original&password=DC66A2AA2B244831ACB43C8B4E290857A27270DB3E554F17BC19CA2D720F5D8A", UNESCAPE("HN004_2024-10-09_%5B26_1766896_5826%5D.JPG"), "IMG20241009_163826.jpg", UNESCAPE("1766896_B%u1EAEC%20GIANG_HN004_2024-10-09_26__IMAGE_PRODUCT.JPG"), "jpg", "11033622", UNESCAPE("HN004_UBND%20TP.Bac%20Giang_B%u1EAEC%20GIANG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11032954&ImageType=original&password=E48ADB4852AD4A0CAFB886E1E24AF5107D3E8335830C40C0B61647EE8AB48361", UNESCAPE("HN009_2024-10-09_%5B27_1766824_5826%5D.JPG"), "IMG20241009_161034.jpg", UNESCAPE("1766824_NGH%u1EC6%20AN_HN009_2024-10-09_27__IMAGE_PRODUCT.JPG"), "jpg", "11032954", UNESCAPE("HN009_UBND%20TP.Vinh_NGH%u1EC6%20AN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030541&ImageType=original&password=44F7DCBC3EFF45FAA4ABCF7392320955F7A4B4C875724C30BBF569C9AB1A51EA", UNESCAPE("HN007_2024-10-09_%5B27_1766577_5826%5D.JPG"), "IMG20241009_162306.jpg", UNESCAPE("1766577_THANH%20H%D3A_HN007_2024-10-09_27__IMAGE_PRODUCT.JPG"), "jpg", "11030541", UNESCAPE("HN007_UBND%20Tp.Thanh%20H%F3a_THANH%20H%D3A"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11032922&ImageType=original&password=C1FB942BDF1B427A99D68FD53CBAF48682C82DF967234386AF75C804E6FD46CB", UNESCAPE("HN005_2024-10-09_%5B26_1766821_5826%5D.JPG"), "IMG20241009_154931.jpg", UNESCAPE("1766821_H%u1EA2I%20D%u01AF%u01A0NG_HN005_2024-10-09_26__IMAGE_PRODUCT.JPG"), "jpg", "11032922", UNESCAPE("HN005_UBND%20TP.Hai%20Duong_H%u1EA2I%20D%u01AF%u01A0NG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033265&ImageType=original&password=2E89070FF5144E77ADA6CCF9C10C8E833F70862139864860AF1B3976F96532BF", UNESCAPE("HN005_2024-10-09_%5B30_1766840_5826%5D.JPG"), "IMG20241009_153442.jpg", UNESCAPE("1766840_H%u1EA2I%20D%u01AF%u01A0NG_HN005_2024-10-09_30__IMAGE_PRODUCT.JPG"), "jpg", "11033265", UNESCAPE("HN005_UBND%20TP.Hai%20Duong_H%u1EA2I%20D%u01AF%u01A0NG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033445&ImageType=original&password=AE4C666D76E94245B76BC03B87B3159CCE3912A0DC3841E08A2FA2177BCDD515", UNESCAPE("HN002_2024-10-09_%5B26_1766872_5826%5D.JPG"), "IMG20241009_142806.jpg", "1766872_HN_HN002_2024-10-09_26__IMAGE_PRODUCT.JPG", "jpg", "11033445", UNESCAPE("HN002_UBND%20Hoang%20Mai_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033342&ImageType=original&password=644D22B1168E4514A614103B0B2BBCB23A5333BCCDBB4ACFA66BB0A904AED45C", UNESCAPE("HN008_2024-10-09_%5B28_1766864_5826%5D.JPG"), "IMG20241009_125625.jpg", UNESCAPE("1766864_H%u1EA2I%20PH%D2NG_HN008_2024-10-09_28__IMAGE_PRODUCT.JPG"), "jpg", "11033342", UNESCAPE("HN008_UBND%20TP.H%u1EA3i%20Ph%F2ng_H%u1EA2I%20PH%D2NG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033163&ImageType=original&password=F1A637FFD1224FF18071200819A679EBD5CF486A050046198A06985622AD8C0E", UNESCAPE("HN008_2024-10-09_%5B26_1766836_5826%5D.JPG"), "IMG20241009_114757.jpg", UNESCAPE("1766836_H%u1EA2I%20PH%D2NG_HN008_2024-10-09_26__IMAGE_PRODUCT.JPG"), "jpg", "11033163", UNESCAPE("HN008_UBND%20TP.H%u1EA3i%20Ph%F2ng_H%u1EA2I%20PH%D2NG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11029300&ImageType=original&password=263598FD5504427687BA3431492F2D3C8804A8B3F620470A8A5943E21EF5BDF0", UNESCAPE("SG003_2024-10-09_%5B26_1766406_5826%5D.JPG"), "IMG_UPLOAD_20241009_111534.jpg", "1766406_SG_SG003_2024-10-09_26__IMAGE_PRODUCT.JPG", "jpg", "11029300", UNESCAPE("SG003_UBND%20G%F2%20V%u1EA5p_SG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030501&ImageType=original&password=05F7FCED1B9443719A41CF5466F1BD933ADDAE0D5023407799D56BB2DD3B1C58", UNESCAPE("HN009_2024-10-09_%5B29_1766575_5826%5D.JPG"), "IMG20241009_105520.jpg", UNESCAPE("1766575_NGH%u1EC6%20AN_HN009_2024-10-09_29__IMAGE_PRODUCT.JPG"), "jpg", "11030501", UNESCAPE("HN009_UBND%20TP.Vinh_NGH%u1EC6%20AN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033563&ImageType=original&password=30C1BC4D946F4C00B833CAB56B2F2FF783CCA0EA2AB646E8949D029FA2927522", UNESCAPE("HN008_2024-10-09_%5B28_1766881_5826%5D.JPG"), "IMG20241009_104759.jpg", UNESCAPE("1766881_H%u1EA2I%20PH%D2NG_HN008_2024-10-09_28__IMAGE_PRODUCT.JPG"), "jpg", "11033563", UNESCAPE("HN008_UBND%20TP.H%u1EA3i%20Ph%F2ng_H%u1EA2I%20PH%D2NG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033416&ImageType=original&password=9685F3DF17B5489A8798E367C9D1E0563B2B5714887843BEB96F57F0EDD34674", UNESCAPE("HN002_2024-10-09_%5B26_1766871_5826%5D.JPG"), "IMG20241009_102050.jpg", "1766871_HN_HN002_2024-10-09_26__IMAGE_PRODUCT.JPG", "jpg", "11033416", UNESCAPE("HN002_UBND%20Hoang%20Mai_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11029464&ImageType=original&password=02D493DFCE314AF7AF0B2563C7CC6E2FFAF29B2414394DADB6C441C30847DEB6", UNESCAPE("HN002_2024-10-09_%5B26_1766452_5826%5D.JPG"), "IMG20241009_101347.jpg", "1766452_HN_HN002_2024-10-09_26__IMAGE_PRODUCT.JPG", "jpg", "11029464", UNESCAPE("HN002_UBND%20Hoang%20Mai_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030602&ImageType=original&password=106A527402B142CEB11FC3D94EE2B0CAA5C22E3B46D84F00AF5C88474034AEA0", UNESCAPE("SG004_2024-10-09_%5B26_1766584_5826%5D.JPG"), "IMG_20241009_090641.jpg", "1766584_SG_SG004_2024-10-09_26__IMAGE_PRODUCT.JPG", "jpg", "11030602", UNESCAPE("SG004_UBND%20B%ECnh%20Th%u1EA1nh_SG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11029945&ImageType=original&password=6539306041314CCFA0EE5EB4CA01A2982EF4AD9FD8FF44EB96D5B7BA9D049336", UNESCAPE("SG002_2024-10-08_%5B25_1766533_5826%5D.JPG"), "IMG_20241008_184734.jpg", "1766533_SG_SG002_2024-10-08_25__IMAGE_PRODUCT.JPG", "jpg", "11029945", UNESCAPE("SG002_332%20%u0110O%C0N%20V%u0102N%20B%u01A0%2C%20P13%2C%20Q4_SG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11029396&ImageType=original&password=5722C7582545403F886481C4EA2348E64A449B38E11C46B18837698F534CB0F4", UNESCAPE("HN001_2024-10-08_%5B27_1766426_5826%5D.JPG"), "IMG20241008_184446.jpg", "1766426_HN_HN001_2024-10-08_27__IMAGE_PRODUCT.JPG", "jpg", "11029396", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11029142&ImageType=original&password=985ACD2BD7D0471E86E27CF9DFA065ECC31CEF82070745AC953A3A1F3D1B4AE0", UNESCAPE("HN001_2024-10-08_%5B26_1766390_5826%5D.JPG"), "IMG20241008_180654.jpg", "1766390_HN_HN001_2024-10-08_26__IMAGE_PRODUCT.JPG", "jpg", "11029142", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11032884&ImageType=original&password=62BEB37B288B46E5A386526A3B8B9F608B28E9952DAE465D865D1EAA27B0BCF3", UNESCAPE("HN009_2024-10-08_%5B26_1766819_5826%5D.JPG"), "IMG20241008_150635.jpg", UNESCAPE("1766819_NGH%u1EC6%20AN_HN009_2024-10-08_26__IMAGE_PRODUCT.JPG"), "jpg", "11032884", UNESCAPE("HN009_UBND%20TP.Vinh_NGH%u1EC6%20AN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030127&ImageType=original&password=BC399AB08B6F485CACF1F0A103D86A7C9E9EBF9908284261919BC1D1D246C2BF", UNESCAPE("HN001_2024-10-08_%5B28_1766554_5826%5D.JPG"), "IMG20241008_150230.jpg", "1766554_HN_HN001_2024-10-08_28__IMAGE_PRODUCT.JPG", "jpg", "11030127", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033533&ImageType=original&password=017C27EC39A6457988D729C26B115E9F336AB2189FCC4F63ABE3C1DED5B50345", UNESCAPE("HN001_2024-10-08_%5B26_1766879_5826%5D.JPG"), "IMG20241008_150027.jpg", "1766879_HN_HN001_2024-10-08_26__IMAGE_PRODUCT.JPG", "jpg", "11033533", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033475&ImageType=original&password=7D54878596874C689ECCA4341C47A89EB7842B6DD7F940E3871F1D2D7B07E3A6", UNESCAPE("HN001_2024-10-08_%5B27_1766875_5826%5D.JPG"), "IMG20241008_141715.jpg", "1766875_HN_HN001_2024-10-08_27__IMAGE_PRODUCT.JPG", "jpg", "11033475", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033504&ImageType=original&password=E2EE963B1B46445489023294061CEEFE498AD34D0D1B426A90759285E1DBE8E6", UNESCAPE("HN001_2024-10-08_%5B26_1766877_5826%5D.JPG"), "IMG20241008_131255.jpg", "1766877_HN_HN001_2024-10-08_26__IMAGE_PRODUCT.JPG", "jpg", "11033504", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033311&ImageType=original&password=98B54EB2937B4CCCB74024393BCDE80EDA08D566F885442299BE69285DE56FCF", UNESCAPE("HN001_2024-10-08_%5B26_1766845_5826%5D.JPG"), "IMG_20241008_122807.jpg", "1766845_HN_HN001_2024-10-08_26__IMAGE_PRODUCT.JPG", "jpg", "11033311", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030689&ImageType=original&password=7D8860C2869547CAAAA930F887CFF4DD098F6CBCECBF49B69B0B2C28628A8F32", UNESCAPE("SG003_2024-10-08_%5B26_1766589_5826%5D.JPG"), "IMG20241008_111512.jpg", "1766589_SG_SG003_2024-10-08_26__IMAGE_PRODUCT.JPG", "jpg", "11030689", UNESCAPE("SG003_UBND%20G%F2%20V%u1EA5p_SG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11029381&ImageType=original&password=8C4D125F5ADD468E8614A4E0A5A09B2862063E35101446D8B6ABBB1103DC009A", UNESCAPE("SG004_2024-10-08_%5B26_1766424_5826%5D.JPG"), "IMG_1728400347109_1728360440772.jpg", "1766424_SG_SG004_2024-10-08_26__IMAGE_PRODUCT.JPG", "jpg", "11029381", UNESCAPE("SG004_UBND%20B%ECnh%20Th%u1EA1nh_SG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033387&ImageType=original&password=7A4BBF3D4BBA486C95842EEC8FEAD85770A84D6D6ABD4AF4B890A90FC28C2F20", UNESCAPE("HN001_2024-10-08_%5B26_1766869_5826%5D.JPG"), "IMG20241008_102533.jpg", "1766869_HN_HN001_2024-10-08_26__IMAGE_PRODUCT.JPG", "jpg", "11033387", UNESCAPE("HN001_UBND%20HBT_HN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11030789&ImageType=original&password=F6D094DF02B346909C371C3B25B7143E8B3BCBDE67AE40C4AD0F70E564839162", UNESCAPE("HN009_2024-10-08_%5B31_1766675_5826%5D.JPG"), "IMG20241008_100514.jpg", UNESCAPE("1766675_NGH%u1EC6%20AN_HN009_2024-10-08_31__IMAGE_PRODUCT.JPG"), "jpg", "11030789", UNESCAPE("HN009_UBND%20TP.Vinh_NGH%u1EC6%20AN"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11029338&ImageType=original&password=873891CE2D5D4E2ABD977E5440AAF691824FF9CF33CB436587F8C7E0F040A405", UNESCAPE("SG004_2024-10-08_%5B26_1766423_5826%5D.JPG"), "IMG_1728398122569_1728355940686.jpg", "1766423_SG_SG004_2024-10-08_26__IMAGE_PRODUCT.JPG", "jpg", "11029338", UNESCAPE("SG004_UBND%20B%ECnh%20Th%u1EA1nh_SG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11032567&ImageType=original&password=6F9624829F9A42999A1504158BD50A841431462BCC2B4B20901A4A53707AB501", UNESCAPE("SG002_2024-10-08_%5B32_1766799_5826%5D.JPG"), "IMG_20241009_194940.jpg", "1766799_SG_SG002_2024-10-08_32__IMAGE_PRODUCT.JPG", "jpg", "11032567", UNESCAPE("SG002_332%20%u0110O%C0N%20V%u0102N%20B%u01A0%2C%20P13%2C%20Q4_SG"))
Call ProcessOneAttachment("https://ifieldVN.ipsos.com/getImage.asp?ID=11033054&ImageType=original&password=DE83307EDA424A26B51D5C04A6E7E30DA71616A84B124395A6412FF27862D068", UNESCAPE("HN009_2024-10-08_%5B69_1766826_5826%5D.JPG"), "IMG20241008_093410.jpg", UNESCAPE("1766826_NGH%u1EC6%20AN_HN009_2024-10-08_69__IMAGE_PRODUCT.JPG"), "jpg", "11033054", UNESCAPE("HN009_UBND%20TP.Vinh_NGH%u1EC6%20AN"))
MsgBox "Done"
Call WriteErrorLog()
Function DownloadAttachment(strUrl)
	Dim oXMLHTTP
	set oXMLHTTP = CreateObject("Msxml2.ServerXMLHTTP")
	oXMLHTTP.setTimeouts 6000,6000,180000,900000
	If Not bFlagDownloadAttachmentsWithFooter Then
		strUrl = Replace(strUrl, "/getImage.asp?ID=", "/getAttachment.asp?ID=")
	End If
	Call oXMLHTTP.Open("GET", strUrl , False, "", "")
	Call oXMLHTTP.Send()
	If oXMLHTTP.status = 204 Then
		strErrLog = strErrLog & "We're sorry, the link or file you are requesting has expired and is no longer available. URL = " & strUrl & vbCrLf
		DownloadAttachment = Null
		Exit Function
	End If
	If Not oXMLHTTP.status = 200 Then
		If InStr(UCASE(strUrl), UCASE("/getImage.asp?ID=")) > 0 Then
			strErrLog = strErrLog & "Error occurred while trying to contact the service provider. WILL TRY WITH GETATTACHMENT. URL = " & strUrl & vbCrLf
			strUrl = Replace(strUrl, "/getImage.asp?ID=", "/getAttachment.asp?ID=")
			set oXMLHTTP = CreateObject("Msxml2.ServerXMLHTTP")
 			oXMLHTTP.setTimeouts 6000,6000,180000,900000
			Call oXMLHTTP.Open("GET", strUrl , False, "", "")
			Call oXMLHTTP.Send()
			If Not oXMLHTTP.status = 200 Then
				strErrLog = strErrLog & "Error occurred while trying to contact the service provider. GETATTACHMENT FAILED. URL = " & strUrl & vbCrLf
				DownloadAttachment = Null
				Exit Function
			End If
		Else
			strErrLog = strErrLog & "Error occurred while trying to contact the service provider. URL = " & strUrl & vbCrLf
			DownloadAttachment = Null
			Exit Function
		End If
	End If
	DownloadAttachment = oXMLHTTP.responseBody
	set oXMLHTTP = Nothing
End Function
Function GetNameTemplate(attID, template, extension)
	Dim oXMLHTTP, strUrl
	set oXMLHTTP = CreateObject("Msxml2.ServerXMLHTTP")
	'attNameTemplate
	If extension <> "" Then
		strUrl = "https://iFieldVN.ipsos.com/mystservices/Attachments/getAttachmentNameByTemplate.asp?ID=" & attID & "&TemplateName=" & URLEncode(template, True)  & "." & extension
	Else 
		strUrl = "https://iFieldVN.ipsos.com/mystservices/Attachments/getAttachmentNameByTemplate.asp?ID=" & attID & "&TemplateName=" & URLEncode(template, True)
	End If
	oXMLHTTP.setTimeouts 6000,6000,180000,900000
	Call oXMLHTTP.Open("GET", strUrl , False, "", "")
	Call oXMLHTTP.Send()
	If Not oXMLHTTP.status = 200 Then
		strErrLog = strErrLog & "Error occurred while trying to contact the service provider. GETATTACHMENT FAILED. URL = " & strUrl & vbCrLf
		GetNameTemplate = ""
		Exit Function
	End If
	GetNameTemplate = REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE(REPLACE(oXMLHTTP.responseText,"\","_"), "/", "_"), ":", "_"), "?" ,"_"), "<", "_"), ">", "_"), "|", "_"), "*", "_"), """", "_")
End Function
Function SaveAttachment(binAttachment, strFileName, strOriginalFileName, strSpecialFileName, extension, attID, FolderName)
	Dim s, fso, CurrentDirectory, downloadDirectory, AttachmentName
	set s = CreateObject("ADODB.Stream")
	s.Open()
	s.Type = 1 ''' 2 - adTypeText  ;   1 - adTypeBinary
	s.Write binAttachment
	IF bFlagDownloadAttachmentsWithOriginalName THEN 
		AttachmentName = strOriginalFileName 
	ELSE 
		IF bFlagDownloadAttachmentsWithSpecialName THEN 
			AttachmentName = strSpecialFileName 
		ELSE 
			IF attNameTemplate <> "" THEN
				AttachmentName = GetNameTemplate(attID, attNameTemplate, extension)
				If AttachmentName = "" Then
					Exit Function
				End If
			ELSE 
				AttachmentName = strFileName 
			END IF 
		END IF 
	END IF 
	
	IF downloadInFolders And FolderName <> "" Then
		If foldersNameTemplate <> "" Then
			FolderName = GetNameTemplate(attID, foldersNameTemplate, "")
			If FolderName = "" Then
				Exit Function
			End If
		End If
		On Error Resume Next
			Set fso = CreateObject("Scripting.FileSystemObject")
			CurrentDirectory = fso.GetAbsolutePathName(".")
			downloadDirectory = fso.BuildPath(CurrentDirectory, FolderName)
			
			If Not fso.FolderExists(downloadDirectory) Then
				fso.CreateFolder(downloadDirectory)
			End If
			
			s.SaveToFile fso.BuildPath(downloadDirectory, AttachmentName), adSaveCreateOverWrite
			
			If Err.Number <> 0 Then
				s.SaveToFile AttachmentName, adSaveCreateOverWrite
			End If
		On Error Goto 0
	ELSE
		s.SaveToFile AttachmentName, adSaveCreateOverWrite
	END IF
	
	s.close
	set s = Nothing
End Function
Function ProcessOneAttachment(strUrl, strFileName, strOriginalFileName, strSpecialFileName, extension, attID, FolderName)
	Dim localAttList: localAttList = ""
	If InStr(localAttList, attID) Then
		strErrLog = strErrLog & "We're sorry, the link or file you are requesting has expired and is no longer available. URL = " & strUrl & vbCrLf
		Exit Function
	End If
	Dim binAttachment
	binAttachment = DownloadAttachment(strUrl)
	strFileName = REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE(REPLACE(strFileName,"\",""), "/", ""), ":", ""), "?" ,""), "<", ""), ">", ""), "|", ""), "*", ""), """", "")
	strOriginalFileName = REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE(REPLACE(strOriginalFileName,"\",""), "/", ""), ":", ""), "?" ,""), "<", ""), ">", ""), "|", ""), "*", ""), """", "")
	strSpecialFileName = REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE(REPLACE(strSpecialFileName,"\",""), "/", ""), ":", ""), "?" ,""), "<", ""), ">", ""), "|", ""), "*", ""), """", "")
	attNameTemplate = REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE(REPLACE(attNameTemplate,"\",""), "/", ""), ":", ""), "?" ,""), "<", ""), ">", ""), "|", ""), "*", ""), """", "")
	FolderName = REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE(REPLACE(FolderName,"\",""), "/", ""), ":", ""), "?" ,""), "<", ""), ">", ""), "|", ""), "*", ""), """", "")
	If Not IsNull(binAttachment) Then Call SaveAttachment(binAttachment, strFileName, strOriginalFileName, strSpecialFileName, extension, attID, FolderName)
End Function
Function WriteErrorLog()
	Dim s
	If strErrLog = "" Then Exit Function
	set s = CreateObject("ADODB.Stream")
	s.Open()
	s.Type = 2 
	s.Charset = "unicode"
	s.WriteText strErrLog & vbCrLf
	s.SaveToFile "errorlog.txt", adSaveCreateOverWrite
	s.close
	set s = Nothing
	strErrLog = ""
End Function
Function URLEncode(StringToEncode, UsePlusRatherThanHexForSpace)
	Dim TempAns, CurChr, iChar
	CurChr = 1
	Do Until CurChr - 1 = Len(StringToEncode)
		iChar = Asc(Mid(StringToEncode, CurChr, 1))
		If (iChar > 47 And iChar < 58)  Or (iChar > 64 And iChar < 91) Or (iChar > 96 And iChar < 123) Then
			TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
		ElseIf iChar = 32 Then
			If UsePlusRatherThanHexForSpace Then
				TempAns = TempAns & "+"
			Else
				TempAns = TempAns & "%" & Hex(32)
			End If
		Else
			TempAns = TempAns & "%" & Right("00" & Hex(Asc(Mid(StringToEncode, CurChr, 1))), 2)
		End If
		CurChr = CurChr + 1
	Loop
	URLEncode = TempAns
End Function
