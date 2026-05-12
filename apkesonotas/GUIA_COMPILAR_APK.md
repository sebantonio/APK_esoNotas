# Guía: Compilar APK esoNotas

## 1. Requisitos previos

- **Android Studio** instalado
- **Node.js** instalado
- **VS Code** con extensión **Gradle for Java** (opcional, para ver tareas Gradle)

---

## 2. Primera vez — preparar el proyecto

```powershell
cd apkesonotas
npm install @capacitor/android --legacy-peer-deps --strict-ssl=false
npx cap add android
```

---

## 3. Problema SSL (red corporativa / Avast)

### Importar certificado Avast en Java de Android Studio
Ejecutar en **PowerShell como Administrador**:

```powershell
# Exportar certificado Avast
$cert = Get-ChildItem Cert:\LocalMachine\Root | Where-Object { $_.Subject -like "*Avast*" }
$bytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
[System.IO.File]::WriteAllBytes("C:\Users\PC\avast-root.cer", $bytes)

# Importar en cacerts de Android Studio
& "C:\Program Files\Android\Android Studio\jbr\bin\keytool.exe" `
  -import -trustcacerts -alias "avast-webshield" `
  -file "C:\Users\PC\avast-root.cer" `
  -keystore "C:\Program Files\Android\Android Studio\jbr\lib\security\cacerts" `
  -storepass changeit -noprompt
```

---

## 4. Problema Java 25 incompatible

Añadir en `android/local.properties`:

```
org.gradle.java.home=C\:\\Program Files\\Android\\Android Studio\\jbr
```

---

## 5. Problema Gradle sin internet

Descargar manualmente desde el navegador:
```
https://services.gradle.org/distributions/gradle-8.14.3-all.zip
```

Colocar en:
```
C:\Users\PC\.gradle\wrapper\dists\gradle-8.14.3-all\10utluxaxniiv4wxiphsi49nj\gradle-8.14.3-all.zip
```

---

## 6. Problema ruta con tildes

Añadir en `android/gradle.properties`:
```
android.overridePathCheck=true
```

---

## 7. Actualizar y compilar

Cada vez que modifiques archivos en `www/`:

```powershell
npx cap sync android
```

Luego en **Android Studio**:

1. Menú superior → `Build`
2. `Generate App Bundles or APKs`
3. `Build APK(s)`
4. Esperar a que la barra inferior diga **"Build finished"**
5. Click en **"locate"** en la notificación (abajo a la derecha)

APK generada en:
```
android/app/build/outputs/apk/debug/app-debug.apk
```

---

## 8. Cómo funciona la app en Android

- El botón **Explorar** abre el selector de archivos nativo del dispositivo.
- Al seleccionar un `.xlsx`, se carga en memoria con `xlsx.js`.
- Al **guardar**, el Excel modificado se descarga automáticamente a la carpeta Descargas.
- El archivo se mantiene entre sesiones via `localStorage`.
