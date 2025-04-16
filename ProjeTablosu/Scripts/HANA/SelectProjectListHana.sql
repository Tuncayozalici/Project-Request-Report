SELECT 
    "DocNum"         AS "Döküman Numarası",
    "U_ProjectTitle" AS "Proje Talebi Tanımı",
    "U_NAME"         AS "Talep Eden Kullanıcı",
    "U_Branch"       AS "Şube",
    "U_Department"   AS "Departman",
    "U_RegDate"      AS "Kayıt Tarihi",
    "U_DelDate"      AS "İstenilen Tarih",
    CASE 
        WHEN "U_IsConverted" = 'Y' THEN 'Evet'
        ELSE 'Hayır'
    END AS "Projeye Dönüştürüldü"
FROM "@PROJECT"
WHERE 1 = 1
--DATEFILTER--
--CONVERTEDFILTER--
