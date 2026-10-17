/* =====================================================================
   PoliMusic Web - Datos de roles y usuarios requeridos por la aplicación
   Ejecutar DESPUÉS de BDD_PoliMusic.sql. Es idempotente (se puede repetir).

   Roles (TBL_CATALOG, padre '1'):  0 = USER, 1 = ADMIN
   Contraseñas: por ahora en texto plano (sin cifrar), como userNormal/userAdmin
   en el script original. La app también acepta el hash SHA-1 de 'username1'.
   ===================================================================== */
USE [BDD_PoliMusic];
GO

-- Catálogo de roles (por si no existe)
IF NOT EXISTS (SELECT 1 FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = '' AND COD_CATALOG = '1')
    INSERT INTO TBL_CATALOG (COD_CATALOG_PARENT, COD_CATALOG, VALUE) VALUES ('', '1', 'USER TYPE');
IF NOT EXISTS (SELECT 1 FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = '1' AND COD_CATALOG = '0')
    INSERT INTO TBL_CATALOG (COD_CATALOG_PARENT, COD_CATALOG, VALUE) VALUES ('1', '0', 'USER');
IF NOT EXISTS (SELECT 1 FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = '1' AND COD_CATALOG = '1')
    INSERT INTO TBL_CATALOG (COD_CATALOG_PARENT, COD_CATALOG, VALUE) VALUES ('1', '1', 'ADMIN');
GO

-- Administrador: userAdmin (contraseña: 12345678)
IF NOT EXISTS (SELECT 1 FROM TBL_USER WHERE USERNAME = 'userAdmin')
    INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE)
    VALUES ('userAdmin', '12345678', 'usrfakeadmin@epn.edu.ec', '2001-12-16', '~/userImages/root.png', 1);
ELSE
    UPDATE TBL_USER SET USER_TYPE = 1 WHERE USERNAME = 'userAdmin';
GO

-- Usuario normal: paul.davila (contraseña: 12345678)
IF NOT EXISTS (SELECT 1 FROM TBL_USER WHERE USERNAME = 'paul.davila')
    INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE)
    VALUES ('paul.davila', '12345678', 'paul.davila@epn.edu.ec', NULL, '~/userImages/default.jpg', 0);
ELSE
    UPDATE TBL_USER SET USER_TYPE = 0, PASSWORD = '12345678' WHERE USERNAME = 'paul.davila';
GO

SELECT u.ID_USER, u.USERNAME, u.EMAIL, c.VALUE AS ROL
FROM TBL_USER u
LEFT JOIN TBL_CATALOG c ON c.COD_CATALOG_PARENT = '1' AND c.COD_CATALOG = CAST(u.USER_TYPE AS varchar(8))
ORDER BY u.USER_TYPE DESC, u.USERNAME;
GO
