# Excel-XML Converter

Una aplicación web moderna y eficiente para convertir archivos Excel (.xlsx, .xls) a formato XML de manera sencilla.

## 🚀 Características

- **Interfaz intuitiva** con área de arrastrar y soltar para cargar archivos
- **Soporte multiplataforma** para archivos Excel modernos (.xlsx) y formatos antiguos (.xls)
- **Mapeo personalizable** de columnas Excel a elementos y atributos XML
- **Vista previa** del resultado XML antes de la descarga
- **Conversión en lote** para procesar múltiples hojas o archivos simultáneamente
- **Plantillas guardables** para configuraciones de mapeo frecuentes
- **Procesamiento local** que garantiza la privacidad de tus datos

## 🛠️ Tecnologías

- Frontend: Angular 18, TypeScript
- Procesamiento: xlsx (SheetJS) para manejo de Excel, bibliotecas XML nativas
- Diseño responsivo compatible con dispositivos móviles y escritorio

## 📋 Requisitos

### Para usuarios
- Navegador web moderno (Chrome, Firefox, Safari, Edge)
- Conexión a internet (solo para cargar la aplicación)

### Para desarrolladores
- Node.js (v16 o superior)
- Angular CLI v18.2.5 o superior
- TypeScript 5.5.2 o superior

## 🔧 Instalación

### Para usuarios
No se requiere instalación. Simplemente accede a la aplicación desde:

```
https://excel-xml.web.app
```

### Para desarrolladores
1. Clona el repositorio
```
git clone https://github.com/tu-usuario/excel-xml-converter.git
cd excel-xml-converter
```

2. Instala las dependencias
```
npm install
```

3. Inicia el servidor de desarrollo
```
ng serve
```

4. Accede a `http://localhost:4200` en tu navegador

## 💡 Uso

1. Accede a la aplicación web
2. Arrastra y suelta tu archivo Excel o haz clic en el área designada para seleccionarlo
3. Configura las opciones de mapeo según tus necesidades:
   - Define la estructura del XML resultante
   - Asigna columnas Excel a elementos XML
   - Establece atributos y valores predeterminados
4. Visualiza la vista previa del XML generado
5. Descarga el archivo XML resultante

## 🛠️ Tecnologías utilizadas

- **Angular 18**: Framework principal del frontend
- **TypeScript 5.5**: Lenguaje de programación tipado
- **Bootstrap 5.3**: Para UI responsiva
- **ngx-toastr**: Para notificaciones
- **xlsx (SheetJS)**: Biblioteca para procesamiento de archivos Excel
- **RxJS**: Para programación reactiva

## ⚙️ Opciones avanzadas

- **Mapeo jerárquico**: Crea XMLs con estructuras anidadas complejas
- **Filtrado de datos**: Convierte solo las filas que cumplan con criterios específicos
- **Transformación de datos**: Aplica funciones de transformación durante la conversión
- **Validación XML**: Verifica el XML resultante contra un esquema XSD

## 🔒 Privacidad y seguridad

Esta aplicación procesa todos los archivos localmente en tu navegador. Ningún dato se envía a servidores externos, garantizando la confidencialidad y seguridad de tu información.

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Si deseas colaborar:

1. Haz fork del repositorio
2. Crea una rama para tu funcionalidad (`git checkout -b feature/nueva-funcionalidad`)
3. Realiza tus cambios y haz commit (`git commit -m 'Añadir nueva funcionalidad'`)
4. Sube tus cambios (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está licenciado bajo [MIT License](LICENSE).

## 📞 Contacto

Para soporte técnico o consultas: elydeveloperhn@gmail.com

---

Desarrollado con ❤️ por Ely Dev
