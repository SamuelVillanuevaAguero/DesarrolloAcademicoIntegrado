/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/javafx/FXMLController.java to edit this template
 */
package vistas;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.FileChooser;
import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import utilerias.general.ControladorGeneral;

public class RespaldoController implements Initializable {

    @FXML
    private Button botonCerrar;
    @FXML
    private Button botonMinimizar;
    @FXML
    private Button botonRegresar;
    @FXML
    private RadioButton opcionImportar;
    @FXML
    private RadioButton opcionExportar;
    @FXML
    private Label rutaSeleccionada;
    @FXML
    private RadioButton opcionZIP;
    @FXML
    private RadioButton opcionRAR;
    @FXML
    private Button botonRutaRespaldo;
    @FXML
    private Button botonAceptar;
    @FXML
    private Label mensajeTipoExport;

    private ToggleGroup grupoOpciones;
    private ToggleGroup grupoFormato;
    private final String CARPETA_BASE = System.getProperty("user.home") + "/Desktop/Gestion_de_Cursos";
    private File ultimaUbicacion = new File(System.getProperty("user.home"));

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        configurarInterfaz();
        configurarGruposRadio();
        configurarListenerFormato();

        botonCerrar.setOnMouseClicked(event -> {
            try {
                cerrarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        botonMinimizar.setOnMouseClicked(event -> {
            minimizarVentana(event);
        });

        botonRegresar.setOnMouseClicked(event -> {
            try {
                regresarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        botonRutaRespaldo.setOnMouseClicked(event -> {
            seleccionarRuta(event);
        });

        botonAceptar.setOnMouseClicked(event -> {
            procesarRespaldo(event);
        });
    }

    private void configurarInterfaz() {
        // Configuración inicial de visibilidad
        mensajeTipoExport.setVisible(false);
        opcionRAR.setVisible(false);
        opcionZIP.setVisible(false);
        rutaSeleccionada.setFocusTraversable(false);
        botonAceptar.setDisable(true);
    }

    private void configurarGruposRadio() {
        // Grupo para importar/exportar
        grupoOpciones = new ToggleGroup();
        opcionImportar.setToggleGroup(grupoOpciones);
        opcionExportar.setToggleGroup(grupoOpciones);

        // Grupo para ZIP/RAR
        grupoFormato = new ToggleGroup();
        opcionZIP.setToggleGroup(grupoFormato);
        opcionRAR.setToggleGroup(grupoFormato);

        // Listener para cambios en la selección
        grupoOpciones.selectedToggleProperty().addListener((observable, oldValue, newValue) -> {
            if (newValue == opcionExportar) {
                mensajeTipoExport.setVisible(true);
                opcionRAR.setVisible(true);
                opcionZIP.setVisible(true);
            } else {
                mensajeTipoExport.setVisible(false);
                opcionRAR.setVisible(false);
                opcionZIP.setVisible(false);
                grupoFormato.selectToggle(null);
            }
            rutaSeleccionada.setText("");
            botonAceptar.setDisable(true);
        });
    }

    private void configurarListenerFormato() {
        grupoFormato.selectedToggleProperty().addListener((observable, oldValue, newValue) -> {
            if (newValue != null && archivoSeleccionado != null) {
                // Obtener la nueva extensión basada en la selección
                String nuevaExtension = opcionZIP.isSelected() ? "zip" : "rar";

                // Obtener la ruta actual y cambiar la extensión
                String rutaActual = archivoSeleccionado.getAbsolutePath();
                String nuevaRuta = cambiarExtension(rutaActual, nuevaExtension);

                // Actualizar el archivo seleccionado y la ruta mostrada
                archivoSeleccionado = new File(nuevaRuta);
                rutaSeleccionada.setText(nuevaRuta);
            }
        });
    }

    private String cambiarExtension(String rutaArchivo, String nuevaExtension) {
        int ultimoPunto = rutaArchivo.lastIndexOf('.');
        if (ultimoPunto != -1) {
            // Si existe una extensión, la reemplazamos
            return rutaArchivo.substring(0, ultimoPunto + 1) + nuevaExtension;
        } else {
            // Si no existe extensión, la agregamos
            return rutaArchivo + "." + nuevaExtension;
        }
    }

    private String generarNombreSecuencial(String directoryPath, String baseName, String extension) {
        File directory = new File(directoryPath);
        if (!directory.exists()) {
            return baseName + "." + extension;
        }

        // Comprobar si existe el archivo base
        File baseFile = new File(directory, baseName + "." + extension);
        if (!baseFile.exists()) {
            return baseName + "." + extension;
        }

        // Buscar el siguiente número disponible
        int counter = 1;
        File file;
        do {
            file = new File(directory, baseName + "_" + counter + "." + extension);
            counter++;
        } while (file.exists());

        return baseName + "_" + (counter - 1) + "." + extension;
    }

    // Método auxiliar para verificar si un archivo existe
    private boolean archivoExiste(String directoryPath, String fileName) {
        File file = new File(directoryPath + File.separator + fileName);
        return file.exists();
    }

    private void procesarRespaldo(MouseEvent event) {
        try {
            if (opcionExportar.isSelected()) {
                realizarExportacion();
            } else {
                realizarImportacion();
            }
            // Limpiar controles después de una operación exitosa
            limpiarControles();

        } catch (Exception e) {
            mostrarAlerta("Error", "Error: " + e.getMessage());
        }
    }

    private void seleccionarRuta(MouseEvent event) {
        if (grupoOpciones.getSelectedToggle() == null) {
            mostrarAlerta("Error", "Debe seleccionar una opción (Importar/Exportar)");
            return;
        }

        if (opcionExportar.isSelected() && grupoFormato.getSelectedToggle() == null) {
            mostrarAlerta("Error", "Debe seleccionar un formato de exportación (ZIP/RAR)");
            return;
        }

        if (opcionExportar.isSelected()) {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Guardar archivo de respaldo");

            // Establecer la ubicación inicial al último directorio usado
            fileChooser.setInitialDirectory(ultimaUbicacion);

            // Configurar el nombre inicial del archivo
            String extension = opcionZIP.isSelected() ? "zip" : "rar";
            String baseFileName = "Gestion_de_Cursos";

            // Generar el nombre secuencial antes de mostrar el FileChooser
            String nombreSecuencial = generarNombreSecuencial(
                    ultimaUbicacion.getAbsolutePath(),
                    baseFileName,
                    extension
            );

            // Establecer el nombre inicial con el número secuencial si es necesario
            fileChooser.setInitialFileName(nombreSecuencial);

            // Configurar los filtros de extensión según la selección
            FileChooser.ExtensionFilter selectedFilter;
            if (opcionZIP.isSelected()) {
                selectedFilter = new FileChooser.ExtensionFilter("Archivo ZIP (*.zip)", "*.zip");
            } else {
                selectedFilter = new FileChooser.ExtensionFilter("Archivo RAR (*.rar)", "*.rar");
            }

            fileChooser.getExtensionFilters().add(selectedFilter);
            fileChooser.getExtensionFilters().add(
                    new FileChooser.ExtensionFilter("Todos los archivos (*.*)", "*.*")
            );

            fileChooser.setSelectedExtensionFilter(selectedFilter);

            File selectedFile = fileChooser.showSaveDialog(null);
            if (selectedFile != null) {
                // Guardar la ubicación seleccionada
                ultimaUbicacion = selectedFile.getParentFile();

                // Ya no necesitamos generar otro nombre secuencial aquí
                // porque el usuario ya vio y eligió el nombre correcto
                /*// Verificar si el archivo existe y generar nombre secuencial si es necesario
                String directory = selectedFile.getParent();
                String fileName = selectedFile.getName();
                String finalFileName = generarNombreSecuencial(directory,
                        fileName.substring(0, fileName.lastIndexOf('.')),
                        extension);
                        
                selectedFile = new File(directory, finalFileName);*/
                rutaSeleccionada.setText(selectedFile.getAbsolutePath());
                archivoSeleccionado = selectedFile;
                botonAceptar.setDisable(false);
            }
        } else {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Seleccionar archivo de respaldo");

            // Establecer la ubicación inicial al último directorio usado
            fileChooser.setInitialDirectory(ultimaUbicacion);

            fileChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Archivos de respaldo", "*.zip", "*.rar"),
                    new FileChooser.ExtensionFilter("Todos los archivos (*.*)", "*.*")
            );

            File selectedFile = fileChooser.showOpenDialog(null);

            if (selectedFile != null) {
                // Guardar la ubicación seleccionada
                ultimaUbicacion = selectedFile.getParentFile();
                rutaSeleccionada.setText(selectedFile.getAbsolutePath());
                archivoSeleccionado = selectedFile;
                botonAceptar.setDisable(false);
            }
        }
    }
// Agregar variable de clase para mantener la referencia al archivo
    private File archivoSeleccionado;

    private void realizarExportacion() throws IOException {
        File carpetaBase = new File(CARPETA_BASE);
        if (!carpetaBase.exists()) {
            throw new IOException("No se encuentra la carpeta Gestion_de_Cursos en el escritorio");
        }

        if (archivoSeleccionado == null) {
            throw new IOException("No se ha seleccionado una ubicación para el respaldo");
        }

        if (opcionZIP.isSelected()) {
            exportarZIP(carpetaBase, archivoSeleccionado.getAbsolutePath());
        } else {
            exportarRAR(carpetaBase, archivoSeleccionado.getAbsolutePath());
        }
        mostrarAlerta("Éxito", "Respaldo creado exitosamente");
    }

    private void realizarImportacion() throws IOException {
        if (archivoSeleccionado == null) {
            throw new IOException("No se ha seleccionado un archivo para importar");
        }

        // Obtener la ruta del escritorio
        String escritorio = System.getProperty("user.home") + File.separator + "Desktop";
        File carpetaDestino = new File(escritorio);

        // Crear un directorio temporal para la validación
        File dirTemp = new File(escritorio + File.separator + "temp_validacion");
        if (!dirTemp.exists()) {
            dirTemp.mkdir();
        }

        try {
            // Extraer temporalmente para validar la estructura
            if (archivoSeleccionado.getName().endsWith(".zip")) {
                importarZIP(archivoSeleccionado, dirTemp);
            } else if (archivoSeleccionado.getName().endsWith(".rar")) {
                importarRAR(archivoSeleccionado, dirTemp);
            } else {
                eliminarDirectorio(dirTemp);
                throw new IOException("El archivo debe ser ZIP o RAR");
            }

            // Validar la estructura del directorio
            if (!validarEstructuraDirectorios(dirTemp)) {
                eliminarDirectorio(dirTemp);
                throw new IOException("El archivo de respaldo debe contener el directorio 'Gestion_de_Cursos' "
                        + "con las carpetas: Archivos_importados, Archivos_exportados y Sistema");
            }

            // Si la validación es exitosa, eliminar el directorio existente
            File directorioExistente = new File(CARPETA_BASE);
            if (directorioExistente.exists()) {
                eliminarDirectorio(directorioExistente);
            }

            // Mover el directorio validado a su ubicación final
            File dirGestionCursos = new File(dirTemp, "Gestion_de_Cursos");
            if (dirGestionCursos.exists()) {
                File destinoFinal = new File(carpetaDestino, "Gestion_de_Cursos");
                if (!dirGestionCursos.renameTo(destinoFinal)) {
                    throw new IOException("No se pudo mover el directorio a su ubicación final");
                }
            }

            mostrarAlerta("Éxito", "Respaldo importado exitosamente");
        } finally {
            // Limpiar el directorio temporal
            if (dirTemp.exists()) {
                eliminarDirectorio(dirTemp);
            }
        }
    }

    private boolean validarEstructuraDirectorios(File directorioBase) {
        File dirGestionCursos = new File(directorioBase, "Gestion_de_Cursos");
        if (!dirGestionCursos.exists() || !dirGestionCursos.isDirectory()) {
            return false;
        }

        // Validar la existencia de las tres carpetas requeridas
        boolean tieneArchivosImportados = new File(dirGestionCursos, "Archivos_importados").exists();
        boolean tieneArchivosExportados = new File(dirGestionCursos, "Archivos_exportados").exists();
        boolean tieneSistema = new File(dirGestionCursos, "Sistema").exists();

        return tieneArchivosImportados && tieneArchivosExportados && tieneSistema;
    }

    private void limpiarControles() {
        // Limpiar selecciones de radio buttons
        grupoOpciones.selectToggle(null);
        grupoFormato.selectToggle(null);

        // Ocultar controles de exportación
        mensajeTipoExport.setVisible(false);
        opcionRAR.setVisible(false);
        opcionZIP.setVisible(false);

        // Limpiar ruta y archivo seleccionado
        rutaSeleccionada.setText("");
        archivoSeleccionado = null;  // Importante: resetear el archivo seleccionado

        // Deshabilitar botón aceptar
        botonAceptar.setDisable(true);
    }

    private void exportarZIP(File origen, String rutaDestino) throws ZipException, IOException {
        try (ZipFile zipFile = new ZipFile(rutaDestino)) {
            zipFile.addFolder(origen);
        }
    }

    private void exportarRAR(File origen, String rutaDestino) throws IOException {
        String winRarPath = buscarWinRAR();
        if (winRarPath == null) {
            throw new IOException("WinRAR no está instalado en el sistema");
        }

        try {
            // Comando para crear un archivo RAR
            ProcessBuilder processBuilder = new ProcessBuilder(
                    winRarPath,
                    "a", // comando para añadir al archivo
                    "-r", // incluir subdirectorios
                    "-ep1", // excluir la ruta base
                    rutaDestino, // archivo destino
                    origen.getAbsolutePath() // carpeta a comprimir
            );

            Process process = processBuilder.start();
            int exitCode = process.waitFor();

            if (exitCode != 0) {
                throw new IOException("Error al crear el archivo RAR. Código de salida: " + exitCode);
            }
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("Proceso interrumpido al crear el archivo RAR", e);
        }
    }

    private String buscarWinRAR() {
        // Rutas comunes donde se instala WinRAR
        String[] posiblesPaths = {
            "C:\\Program Files\\WinRAR\\WinRAR.exe",
            "C:\\Program Files (x86)\\WinRAR\\WinRAR.exe"
        };

        for (String path : posiblesPaths) {
            if (new File(path).exists()) {
                return path;
            }
        }
        return null;
    }

    private void importarZIP(File archivo, File destino) throws IOException {
        try {
            try (ZipFile zipFile = new ZipFile(archivo)) {
                zipFile.extractAll(destino.getAbsolutePath());
            }
        } catch (Exception e) {
            throw new IOException("Error al extraer el archivo ZIP: " + e.getMessage());
        }
    }

    private void importarRAR(File archivo, File destino) throws IOException {
        String winRarPath = buscarWinRAR();
        if (winRarPath == null) {
            throw new IOException("WinRAR no está instalado en el sistema");
        }

        try {
            // Comando para extraer un archivo RAR
            ProcessBuilder processBuilder = new ProcessBuilder(
                    winRarPath,
                    "x", // comando para extraer
                    "-y", // responder "sí" a todo
                    archivo.getAbsolutePath(), // archivo a extraer
                    destino.getAbsolutePath() // directorio destino
            );

            Process process = processBuilder.start();
            int exitCode = process.waitFor();

            if (exitCode != 0) {
                throw new IOException("Error al extraer el archivo RAR. Código de salida: " + exitCode);
            }
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("Proceso interrumpido al extraer el archivo RAR", e);
        }
    }

    private void eliminarDirectorio(File directorio) throws IOException {
        if (directorio.exists()) {
            File[] archivos = directorio.listFiles();
            if (archivos != null) {
                for (File archivo : archivos) {
                    if (archivo.isDirectory()) {
                        eliminarDirectorio(archivo);
                    } else {
                        if (!archivo.delete()) {
                            throw new IOException("No se pudo eliminar el archivo: " + archivo.getAbsolutePath());
                        }
                    }
                }
            }
            if (!directorio.delete()) {
                throw new IOException("No se pudo eliminar el directorio: " + directorio.getAbsolutePath());
            }
        }
    }

    private void mostrarAlerta(String titulo, String mensaje) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(titulo);
        alert.setHeaderText(null);
        alert.setContentText(mensaje);
        alert.showAndWait();
    }

    public void cerrarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión?", getClass());
    }

    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    public void regresarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.regresar(event, "Principal", getClass());
    }
}
