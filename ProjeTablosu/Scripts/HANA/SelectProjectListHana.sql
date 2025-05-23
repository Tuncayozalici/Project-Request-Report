﻿SELECT 
    "DocNum"           AS "Döküman Numarası",
    "U_ProjectTitle"   AS "Proje Talebi Tanımı",
    "U_NAME"           AS "Talep Eden Kullanıcı",
    "U_Branch"         AS "Şube",
    "U_Department"     AS "Departman",
    "U_RegDate"        AS "Kayıt Tarihi",
    "U_DelDate"        AS "İstenilen Tarih",
    CASE 
        WHEN "U_IsConverted" = 'Y' THEN 'Onaylandı'
        WHEN "U_IsConverted" = 'N' THEN 'Reddedildi'
        WHEN "U_IsConverted" = 'P' THEN 'Beklemede'
        ELSE 'Bilinmiyor'
    END AS "Durum"
FROM "@PROJECT"
WHERE 1 = 1
--DATEFILTER--
--CONVERTEDFILTER--
