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
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.input.MouseEvent;
import utilerias.general.ControladorGeneral;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.aspose.words.Document;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveOptions;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * FXML Controller class
 */
public class ExportacionReconocimientosController implements Initializable {

    // Definimos las constantes para los tipos de celda
    public static final int CELL_TYPE_STRING = 1;
    public static final int CELL_TYPE_NUMERIC = 0;
    public static final int CELL_TYPE_BOOLEAN = 4;
    public static final int CELL_TYPE_FORMULA = 2;

    @FXML
    private Button botonCerrar;
    @FXML
    private Button botonMinimizar;
    @FXML
    private Button botonRegresar;
    @FXML
    private Button botonBuscar;
    @FXML
    private Button botonModificar;
    @FXML
    private Button botonGuardar;
    @FXML
    private TextField txtcodigodelcurso;
    @FXML
    private TextField txtNombreInstructor;
    @FXML
    private TextArea txtFechaCurso;
    @FXML
    private ComboBox<String> txtFormatos;
    @FXML
    private Button botonLimpiar;
    @FXML
    private ComboBox<String> txtHoras;
    @FXML
    private TextArea txtAreaCompetencias;
    @FXML
    private TextArea txtAreaNombreCurso;
    @FXML
    private Label buttonRedireccionar;
    @FXML
    private RadioButton radiobotonCodigo;
    @FXML
    private RadioButton radiobotonNombre;
    @FXML
    private ComboBox<String> checkCodigos;
    @FXML
    private ComboBox<String> checkNombres;

    // Métodos de los botones de la barra superior
    public void cerrarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión?", getClass());
    }

    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    public void regresarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.regresar(event, "Principal", getClass());
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {

        // Configuración inicial de eventos de los botones de la barra superior
        buttonRedireccionar.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.regresar(event, "ImportacionArchivos", getClass());
            } catch (IOException ex) {
                Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        botonCerrar.setOnMouseClicked(event -> {
            try {
                cerrarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(ExportacionReconocimientosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        botonMinimizar.setOnMouseClicked(event -> minimizarVentana(event));

        botonRegresar.setOnMouseClicked(event -> {
            try {
                regresarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(ExportacionReconocimientosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        // Configurar opciones en ComboBox de horas y formatos
        for (int i = 20; i <= 50; i++) {
            txtHoras.getItems().add(String.valueOf(i));
        }
        txtFormatos.getItems().addAll("AMBOS", "WORD", "PDF");

        // Deshabilitar campos al inicio
        txtAreaNombreCurso.setDisable(true);
        txtFechaCurso.setDisable(true);
        txtHoras.setDisable(true);
        txtNombreInstructor.setDisable(true);
        txtAreaCompetencias.setDisable(true);
        txtFormatos.setDisable(true);
        txtcodigodelcurso.setDisable(true);

        checkCodigos.setDisable(true);
        checkNombres.setDisable(true);

        // oculta los txt
        txtcodigodelcurso.setVisible(false); // Oculta el TextField
        txtAreaNombreCurso.setVisible(false);

        // Crear un ToggleGroup y asignarlo a los RadioButton
        ToggleGroup toggleGroup = new ToggleGroup();
        radiobotonCodigo.setToggleGroup(toggleGroup);
        radiobotonNombre.setToggleGroup(toggleGroup);

        checkCodigos.setOnAction(event -> {
            String seleccion = checkCodigos.getValue();
            if (seleccion != null && !seleccion.trim().isEmpty()) {
                txtcodigodelcurso.setText(seleccion);
                txtcodigodelcurso.setVisible(true);
                checkCodigos.setVisible(false);
                txtAreaNombreCurso.clear(); // Limpia el campo de nombres para evitar interferencias
            }
        });

        checkNombres.setOnAction(event -> {
            String seleccion = checkNombres.getValue();
            if (seleccion != null && !seleccion.trim().isEmpty()) {
                txtAreaNombreCurso.setText(seleccion);
                txtAreaNombreCurso.setVisible(true);
                checkNombres.setVisible(false);
                txtcodigodelcurso.clear(); // Limpia el campo de códigos para evitar interferencias
            }
        });

        // Asociar los eventos de los RadioButtons
        radiobotonCodigo.setOnAction(event -> {
            try {
                manejarRadioButtonCodigo();
                checkCodigos.setVisible(true);
                txtcodigodelcurso.setVisible(false);
                txtAreaNombreCurso.clear();
            } catch (IOException ex) {
                Logger.getLogger(ExportacionReconocimientosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        radiobotonNombre.setOnAction(event -> {
            try {
                manejarRadioButtonNombre();
                checkNombres.setVisible(true);
                txtAreaNombreCurso.setVisible(false);
                txtcodigodelcurso.clear();
            } catch (IOException ex) {
                Logger.getLogger(ExportacionReconocimientosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        // Evento para checkCodigos (cuando se selecciona un ítem)
        checkCodigos.setOnAction(event -> {
            if (checkCodigos.getValue() != null) {
                // Asignar el valor seleccionado al campo de texto
                txtcodigodelcurso.setText(checkCodigos.getValue());
                // Acción cuando se ha seleccionado un valor en el ComboBox
                checkCodigos.setVisible(false);
                txtcodigodelcurso.setVisible(true);
            }
        });

        // Evento para checkNombres (cuando se selecciona un ítem)
        checkNombres.setOnAction(event -> {
            if (checkNombres.getValue() != null) {
                // Asignar el valor seleccionado al campo de texto
                txtAreaNombreCurso.setText(checkNombres.getValue());
                // Acción cuando se ha seleccionado un valor en el ComboBox
                checkNombres.setVisible(false);  // Ocultar el CheckBox de nombres cuando algo es seleccionado
                txtAreaNombreCurso.setVisible(true);
            }
        });

    }

    private void manejarRadioButtonCodigo() throws IOException {
        checkCodigos.setDisable(false);
        checkNombres.setDisable(true);

        checkNombres.setVisible(false);
        txtAreaNombreCurso.setVisible(true);

        // Limpiar campos y ComboBox
        txtcodigodelcurso.clear();
        checkCodigos.getItems().clear();
        txtAreaNombreCurso.clear();
        checkNombres.getItems().clear();
        txtNombreInstructor.clear();
        txtAreaCompetencias.clear();
        txtFechaCurso.clear();

        // Cargar códigos en el ComboBox
        cargarDatosComboBox(txtcodigodelcurso.getText());
    }

    private void manejarRadioButtonNombre() throws IOException {
        checkNombres.setDisable(false);
        checkCodigos.setDisable(true);

        checkCodigos.setVisible(false);
        txtcodigodelcurso.setVisible(true);

        // Limpiar campos y ComboBox
        txtAreaNombreCurso.clear();
        checkNombres.getItems().clear();
        txtcodigodelcurso.clear();
        txtNombreInstructor.clear();
        txtAreaCompetencias.clear();
        txtFechaCurso.clear();

        // Cargar nombres en el ComboBox
        cargarDatosComboBox(txtAreaNombreCurso.getText());
    }

    @FXML
    private void buscarCurso(ActionEvent event) {
        txtFormatos.setDisable(false); // Habilitar campo formatos
        txtcodigodelcurso.setDisable(true); // Deshabilitar el campo de código para evitar cambios manuales
        ExcelReader excelReader = new ExcelReader(); // Instancia de ExcelReader para manejar el archivo

        try {
            String nombreCurso = null;
            String codigoCurso = null;

            // Verificar si se seleccionó algo en checkCodigos (buscar por código)
            if (checkCodigos.getValue() != null && !checkCodigos.getValue().isEmpty()) {
                codigoCurso = checkCodigos.getValue();
                nombreCurso = excelReader.buscarCurso(codigoCurso); // Buscar nombre por código
            } // Verificar si se seleccionó algo en checkNombres (buscar por nombre)
            else if (checkNombres.getValue() != null && !checkNombres.getValue().isEmpty()) {
                nombreCurso = checkNombres.getValue(); // Obtener el nombre del curso seleccionado
                // Buscar el código correspondiente al nombre
                codigoCurso = excelReader.obtenerCodigoPorNombre(nombreCurso); // Obtener el código del curso por nombre
            } else {
                Alert alert = new Alert(Alert.AlertType.INFORMATION, "Selecciona un curso.");
                alert.showAndWait();
                txtFormatos.setDisable(true);
                return;
            }

            if (nombreCurso != null && !nombreCurso.isEmpty()) {
                // Si se encontró el curso, llenar los campos correspondientes
                txtAreaNombreCurso.setText(nombreCurso);
                txtcodigodelcurso.setText(codigoCurso); // Llenar el campo con el código

                // Buscar los detalles del curso (fecha, horas, instructor, competencias)
                Map<String, String> datosCurso = excelReader.buscarDetallesCurso(nombreCurso);
                if (datosCurso != null) {
                    // Llenar los campos con los datos encontrados
                    txtFechaCurso.setText(datosCurso.get("fechaCurso"));
                    txtHoras.setValue(datosCurso.get("horasCurso"));
                    txtNombreInstructor.setText(datosCurso.get("nombreInstructor"));
                    txtAreaCompetencias.setText(datosCurso.get("competencias"));
                } else {
                    // Si no se encuentran los detalles, mostrar un mensaje informativo
                    Alert alert = new Alert(Alert.AlertType.INFORMATION, "No se encontraron detalles adicionales para el curso.");
                    alert.showAndWait();
                }
            } else {
                // Si no se encuentra el curso, mostrar un mensaje informativo
                Alert alert = new Alert(Alert.AlertType.INFORMATION, "No se encontró ningún curso con ese código o nombre.");
                alert.showAndWait();
                txtcodigodelcurso.setDisable(false);
            }
        } catch (IOException e) {
            // Manejo de excepciones en caso de error con el archivo
            Alert alert = new Alert(Alert.AlertType.ERROR, "Error al leer el archivo de Excel: " + e.getMessage());
            alert.showAndWait();
            e.printStackTrace();
        }
    }

    @FXML
    private void modificarDatos(ActionEvent event) {
        // Habilitar campos para edición
        txtFechaCurso.setDisable(false);
        txtNombreInstructor.setDisable(false);
        txtAreaCompetencias.setDisable(false);
        txtHoras.setDisable(false);

        // Mantener deshabilitados los campos no editables
        txtAreaNombreCurso.setDisable(true);
        txtcodigodelcurso.setDisable(true);
    }

    public int obtenerUltimaSemana(String carpetaDestino, String nombreArchivo, String versionS, String extension) {
        File carpeta = new File(carpetaDestino);

        // Validar que la carpeta existe y es un directorio
        if (!carpeta.exists() || !carpeta.isDirectory()) {
            Logger.getLogger(this.getClass().getName()).log(Level.WARNING,
                    "La ruta especificada no es válida: " + carpetaDestino);
            Alert alert = new Alert(Alert.AlertType.WARNING);
            alert.setTitle("Ruta");
            alert.setHeaderText(null);
            alert.setContentText("Parece que el periodo no tiene información");
            alert.showAndWait();
            return 1; // Si no es un directorio válido, asumimos que es la primera versión
        }

        // Filtrar archivos que coincidan con el patrón "condensado_(version_X).xlsx"
        File[] archivos = carpeta.listFiles((dir, name) -> name.matches(nombreArchivo));

        if (archivos == null || archivos.length == 0) {
            return 1; // Si no hay archivos, retornamos 1 como primera versión
        }

        // Determinar la versión más alta
        int maxVersion = 0;
        Pattern pattern = Pattern.compile(".*\\(" + versionS + "_(\\d+)\\)\\." + extension + "$"); // Patrón para extraer el número de versión

        for (File archivo : archivos) {
            String nombre = archivo.getName();
            Matcher matcher = pattern.matcher(nombre);
            if (matcher.matches()) {
                try {
                    int version = Integer.parseInt(matcher.group(1)); // Extraer y parsear el número de versión
                    maxVersion = Math.max(maxVersion, version); // Comparar con la versión más alta encontrada
                } catch (NumberFormatException e) {
                    Logger.getLogger(this.getClass().getName()).log(Level.WARNING,
                            "No se pudo parsear el número de versión en el archivo: " + nombre, e);
                }
            }
        }

        return maxVersion; // Retornar la siguiente versión disponible
    }

    @FXML
    private void exportarReconocimientos(ActionEvent event) throws IOException, Exception {
        ExcelReader excelReader = new ExcelReader();

        // Obtener el curso seleccionado desde la interfaz
        String nombreCurso = txtAreaNombreCurso.getText(); // Asumiendo que aquí está el nombre del curso seleccionado

        // Validar que el curso esté seleccionado
        if (nombreCurso == null || nombreCurso.isEmpty()) {
            Alert alerta = new Alert(Alert.AlertType.WARNING, "Por favor, selecciona un curso antes de exportar.");
            alerta.showAndWait();
            return;
        }

        // Obtener la lista de nombres completos filtrados por curso
        List<String> nombresDocentesAcreditados = excelReader.obtenerDocentesAcreditadosPorCurso(nombreCurso);

        // Validar que haya docentes para exportar
        if (nombresDocentesAcreditados.isEmpty()) {
            Alert alerta = new Alert(Alert.AlertType.INFORMATION, "No hay docentes acreditados para este curso.");
            alerta.showAndWait();
            return;
        }

        // Ruta de la plantilla Word y el directorio de salida
        Calendar calendario = Calendar.getInstance();
        int año = calendario.get(Calendar.YEAR);
        int periodo = calendario.get(Calendar.MONTH) < 7 ? 1 : 2;

        String rutaPlantilla = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_importados\\" + año + "\\" + periodo + "-" + año + "\\formato_de_hojas_membretadas_para_reconocimientos\\";

        int versionPlantilla = obtenerUltimaSemana(rutaPlantilla, "formato\\_\\(Version_\\d+\\)\\.docx", "Version", "docx");
        rutaPlantilla += "formato_(Version_" + versionPlantilla + ").docx";

        String directorioBase = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_exportados\\" + año + "\\" + periodo + "-" + año + "\\reconocimientos\\";

        // Crear subdirectorio con el nombre del curso
        String directorioCurso = directorioBase + nombreCurso.trim().replace(" ", "_") + "\\";
        File dirCurso = new File(directorioCurso);
        if (!dirCurso.exists()) {
            dirCurso.mkdirs();
        }
        String horasCurso = txtHoras.getValue(); // Asumiendo que seleccionas las horas desde un ComboBox

        // Verificar el formato seleccionado
        String formatoSeleccionado = txtFormatos.getValue();

        // Validar campos esenciales
        if (horasCurso == null || formatoSeleccionado == null) {
            Alert alerta = new Alert(Alert.AlertType.WARNING, "Por favor, asegúrate de que los datos del curso y el formato estén completos.");
            alerta.showAndWait();
            return;
        }
        // Contador para los reconocimientos exitosamente generados
        int totalExportados = 0;
        // Generar reconocimientos para cada docente
        for (String nombreDocente : nombresDocentesAcreditados) {
            try {
                if (formatoSeleccionado.equals("AMBOS")) {
                    // Generar PDF
                    String archivoPDFGenerado = generarDocumentoPDF(rutaPlantilla, directorioCurso, nombreDocente, horasCurso);
                    totalExportados++;
                } else if (formatoSeleccionado.equals("WORD")) {
                    // Generar documento Word
                    String archivoWordGenerado = generarDocumentoWord(rutaPlantilla, directorioCurso, nombreDocente, horasCurso);
                    totalExportados++;
                } else if (formatoSeleccionado.equals("PDF")) {
                    // Generar tanto PDF como Word
                    generarDocumentoPDF(directorioCurso, generarDocumentoWordEnMemoria(rutaPlantilla, nombreDocente, horasCurso), nombreDocente);
                    //String archivoWordGenerado = generarDocumentoWord(rutaPlantilla, directorioCurso, nombreDocente, horasCurso);
                    totalExportados++;
                }
            } catch (IOException e) {
                e.printStackTrace();
                Alert error = new Alert(Alert.AlertType.ERROR, "Error al generar el reconocimiento para " + nombreDocente + ": " + e.getMessage());
                error.showAndWait();
            }
        }

        // Mostrar mensaje con la cantidad de reconocimientos generados
        Alert exito = new Alert(Alert.AlertType.INFORMATION, "¡Exportación completada!\n"
                + "Reconocimientos generados exitosamente: " + totalExportados);
        exito.showAndWait();
        limpiarCampos(event);
    }

    @FXML
    private void guardarDatos(ActionEvent event) {
        // Crear alerta de confirmación
        Alert confirmacion = new Alert(AlertType.CONFIRMATION);
        confirmacion.setTitle("Confirmación de guardado");
        confirmacion.setHeaderText("¿Está seguro de guardar los cambios?");

        // Mostrar la alerta y esperar respuesta del usuario
        confirmacion.showAndWait().ifPresent(response -> {
            if (response == ButtonType.OK) {
                String nombreCurso = txtAreaNombreCurso.getText();
                Map<String, String> datosActualizados = new HashMap<>();
                datosActualizados.put("competencias", txtAreaCompetencias.getText());
                datosActualizados.put("fechaCurso", txtFechaCurso.getText());
                datosActualizados.put("horasCurso", txtHoras.getValue());
                datosActualizados.put("nombreInstructor", txtNombreInstructor.getText());

                ExcelReader excelReader = new ExcelReader();
                try {
                    if (excelReader.guardarDetallesCurso(nombreCurso, datosActualizados)) {
                        // Mostrar alerta de éxito
                        Alert exito = new Alert(AlertType.INFORMATION);
                        exito.setTitle("Éxito");
                        exito.setHeaderText(null);
                        exito.setContentText("Los cambios se guardaron con éxito.");
                        exito.showAndWait();

                        // Deshabilitar nuevamente los campos
                        txtFechaCurso.setDisable(true);
                        txtNombreInstructor.setDisable(true);
                        txtAreaCompetencias.setDisable(true);
                        txtHoras.setDisable(true);
                    } else {
                        Alert error = new Alert(AlertType.ERROR, "No se pudo encontrar el curso para actualizar.");
                        error.showAndWait();
                    }
                } catch (IOException e) {
                    Alert error = new Alert(AlertType.ERROR, "Error al guardar los datos: " + e.getMessage());
                    error.showAndWait();
                    e.printStackTrace();
                }
            } else {
                System.out.println("Guardado cancelado.");
            }
        });
    }

    @FXML
    private void limpiarCampos(ActionEvent event) {
        // Limpiar todos los TextFields
        txtcodigodelcurso.clear();
        txtAreaNombreCurso.clear();
        txtNombreInstructor.clear();
        txtAreaCompetencias.clear();
        txtFechaCurso.clear();

        // Restablecer ComboBoxes
        txtHoras.getSelectionModel().clearSelection();
        txtFormatos.getSelectionModel().clearSelection();

        //
        checkCodigos.setDisable(true);
        checkNombres.setDisable(true);

        // Deshabilitar campos al inicio
        txtAreaNombreCurso.setDisable(true);
        txtFechaCurso.setDisable(true);
        txtHoras.setDisable(true);
        txtFormatos.setDisable(true);
        txtNombreInstructor.setDisable(true);
        txtAreaCompetencias.setDisable(true);
        txtcodigodelcurso.setDisable(true);
        txtAreaNombreCurso.setDisable(true);

        // limpiar
        checkCodigos.getSelectionModel().clearSelection();
        checkNombres.getSelectionModel().clearSelection();

        // activar combo
        checkCodigos.setVisible(true);
        checkNombres.setVisible(true);

        // ocultar textfields
        txtcodigodelcurso.setVisible(false);
        txtAreaNombreCurso.setVisible(false);

        // Limpiar radioButtones
        radiobotonCodigo.setSelected(false);
        radiobotonNombre.setSelected(false);

    }

    @FXML
    private void seleccionCodigo(ActionEvent event) {
    }

    @FXML
    private void seleccionNombre(ActionEvent event) {
    }

    private String obtenerDirector(String ruta, int año, int periodo) {
        try {
            // Abrir el archivo Excel
            FileInputStream file = new FileInputStream(ruta);
            Workbook libro = new XSSFWorkbook(file);

            // Buscar la hoja por nombre basado en el año
            Sheet hoja = libro.getSheet(String.valueOf(año));
            if (hoja == null) {
                System.err.println("No se encontró una hoja con el nombre: " + año);
                return null; // Salir si no se encuentra la hoja
            }

            // Leer el valor de la celda en la fila 3, columna 2 (Indexada desde 0)
            return hoja.getRow(0).getCell(1).getStringCellValue();

        } catch (FileNotFoundException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Archivo no encontrado: " + ruta, ex);
        } catch (IOException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error al leer el archivo Excel", ex);
        } catch (NullPointerException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error: Celda o fila no encontrada", ex);
        } catch (NumberFormatException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error al convertir el valor de la celda a número", ex);
        }
        return null; // Devolver 0 si ocurre algún error
    }

    @FXML
    private void RedireccionarArchivos(ActionEvent event) {
        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("ImportacionArchivos.fxml"));
            Parent root = loader.load();

            Stage stage = (Stage) buttonRedireccionar.getScene().getWindow(); // Obtener la ventana actual
            stage.setScene(new Scene(root)); // Mostrar la nueva escena
            stage.show();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private String generarDocumentoWord(String rutaPlantilla, String directorioSalida, String nombreDocente, String horasCurso) throws IOException {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XWPFDocument documento = null;

        try {
            // Abrir la plantilla
            fis = new FileInputStream(rutaPlantilla);
            documento = new XWPFDocument(fis);

            // Procesar todos los párrafos del documento
            procesarParrafos(documento, nombreDocente, horasCurso);

            // Guardar el archivo generado
            String rutaArchivoGenerado = directorioSalida + "Reconocimiento_" + nombreDocente + ".docx";
            fos = new FileOutputStream(rutaArchivoGenerado);
            documento.write(fos);

            return rutaArchivoGenerado; // Devolver la ruta del archivo generado
        } finally {
            // Cerrar recursos manualmente
            if (fos != null) {
                fos.close();
            }
            if (fis != null) {
                fis.close();
            }
        }
    }

    public String generarDocumentoPDF(String rutaPlantilla, String directorioSalida, String nombreDocente, String horasCurso) throws Exception {
        // Generar el documento Word temporal
        String rutaWord = generarDocumentoWord(rutaPlantilla, directorioSalida, nombreDocente, horasCurso);

        // Verificar que Aspose pueda leer el documento generado
        File wordFile = new File(rutaWord);
        if (!wordFile.exists()) {
            throw new FileNotFoundException("El archivo Word no se generó correctamente.");
        }

        // Convertir el documento Word a PDF
        String rutaPDF = directorioSalida + "Reconocimiento_" + nombreDocente + ".pdf";

        // Crear el documento Aspose a partir del archivo Word
        Document asposeDoc = new Document(rutaWord);

        // Guardar en formato PDF  
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // Si deseas embebeder las fuentes completamente, puedes probar con esta opción 
        saveOptions.setEmbedFullFonts(true);
        asposeDoc.save(rutaPDF, saveOptions);
        //asposeDoc.save(rutaPDF, com.aspose.words.SaveFormat.PDF);

        return rutaPDF; // Devolver la ruta del archivo PDF generado
    }

    //MÉTODO DE WORD TEMPORAL
    public ByteArrayOutputStream generarDocumentoWordEnMemoria(String rutaPlantilla, String nombreDocente, String horasCurso) throws IOException {
        FileInputStream fis = null;
        org.apache.poi.xwpf.usermodel.XWPFDocument documento = null;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();

        try {
            // Abrir la plantilla
            fis = new FileInputStream(rutaPlantilla);
            documento = new org.apache.poi.xwpf.usermodel.XWPFDocument(fis);

            // Procesar todos los párrafos del documento
            procesarParrafos(documento, nombreDocente, horasCurso);

            // Escribir el documento en memoria
            documento.write(bos);
            return bos;

        } finally {
            if (fis != null) {
                fis.close();
            }
        }
    }

    //MÉTODO DE PRUEBA PDF
    public String generarDocumentoPDF(String directorioSalida, ByteArrayOutputStream documentoWordMemoria, String nombreDocente) throws Exception {
        // Crear un InputStream a partir del documento en memoria
        InputStream inputStream = new ByteArrayInputStream(documentoWordMemoria.toByteArray());

        // Crear el documento Aspose a partir del InputStream
        Document asposeDoc = new Document(inputStream);

        // Ruta del PDF a generar
        String rutaPDF = directorioSalida + "Reconocimiento_" + nombreDocente + ".pdf";

        // Guardar como PDF
        com.aspose.words.PdfSaveOptions saveOptions = new com.aspose.words.PdfSaveOptions();
        saveOptions.setEmbedFullFonts(true);
        asposeDoc.save(rutaPDF, saveOptions);

        return rutaPDF; // Devolver la ruta del PDF generado
    }

    // Procesa todos los párrafos en un documento
    private void procesarParrafos(XWPFDocument documento, String nombreDocente, String horasCurso) {
        // Obtén los valores adicionales de los TextFields y TextAreas
        String codigoCurso = txtcodigodelcurso.getText();
        String nombreCurso = txtAreaNombreCurso.getText();
        String fechaCurso = txtFechaCurso.getText();
        Calendar calendario = Calendar.getInstance();
        int year = calendario.get(Calendar.YEAR);
        int mesActual = calendario.get(Calendar.MONTH) + 1;
        String nombreDirector = obtenerDirector(ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx", year, (mesActual >= 1 && mesActual <= 7) ? 1 : 2);

        for (XWPFParagraph parrafo : documento.getParagraphs()) {
            reemplazarMarcadores(parrafo, nombreDocente, horasCurso, codigoCurso, nombreCurso, fechaCurso, nombreDirector);
        }
    }

    private void reemplazarMarcadores(XWPFParagraph parrafo, String nombreDocente, String horasCurso, String codigoCurso, String nombreCurso, String fechaCurso, String nombreDirector) {
        // Convierte los valores a mayúsculas
        nombreDocente = nombreDocente.toUpperCase();
        nombreDirector = nombreDirector.toUpperCase();
        horasCurso = ("CON UNA DURACIÓN DE " + horasCurso).toUpperCase() + " HORAS";
        codigoCurso = codigoCurso.toUpperCase();
        nombreCurso = nombreCurso.toUpperCase();
        fechaCurso = ("REALIZADO " + fechaCurso).toUpperCase();

        // Obtén el texto completo del párrafo
        String textoCompleto = parrafo.getText();

        // Si contiene algún marcador, realiza el reemplazo
        if (textoCompleto.contains("{nombreDocente}") || textoCompleto.contains("{horasCurso}")
                || textoCompleto.contains("{codigoCurso}") || textoCompleto.contains("{nombreCurso}")
                || textoCompleto.contains("{fechaCurso}") || textoCompleto.contains("{nombreDirector}")) {

            // Unificar texto: primero elimina todos los runs
            for (int i = parrafo.getRuns().size() - 1; i >= 0; i--) {
                parrafo.removeRun(i);
            }

            // Reemplaza todos los marcadores con su valor correspondiente
            textoCompleto = textoCompleto
                    .replace("{nombreDocente}", nombreDocente)
                    .replace("{horasCurso}", horasCurso)
                    .replace("{codigoCurso}", codigoCurso)
                    .replace("{nombreCurso}", nombreCurso)
                    .replace("{fechaCurso}", fechaCurso)
                    .replace("{nombreDirector}", nombreDirector);

            // Crear un solo run para todo el texto (con formato base uniforme)
            XWPFRun runTexto = parrafo.createRun();
            runTexto.setText(textoCompleto);
            runTexto.setFontFamily("Montserrat"); // Forzar la fuente en todo el texto
            runTexto.setFontSize(11);             // Tamaño base uniforme
            runTexto.setColor("595959");          // Color base uniforme
            runTexto.setBold(false);              // Sin negrita por defecto

            // Aplicar formato específico a marcadores clave
            aplicarEstilo(parrafo, nombreCurso, "Montserrat", 18, "595959", true);
            aplicarEstilo(parrafo, nombreDirector, "Montserrat", 12, "595959", true);
            aplicarEstilo(parrafo, nombreDocente, "Montserrat", 24, "595959", true);
        }
    }

    private void aplicarEstilo(XWPFParagraph parrafo, String texto, String fuente, int tamanio, String color, boolean negrita) {
        for (XWPFRun run : parrafo.getRuns()) {
            String textoRun = run.text();
            if (textoRun != null && textoRun.contains(texto)) {
                // Unificar el texto dentro del run para evitar divisiones
                String nuevoTexto = textoRun.replace(texto, texto);
                run.setText(nuevoTexto, 0); // Actualiza todo el texto del run

                // Configurar estilo específico para este texto
                run.setFontFamily(fuente);  // Forzar siempre la fuente
                run.setFontSize(tamanio);
                run.setBold(negrita);
                run.setColor(color);
            }
        }
    }

    private void cargarDatosComboBox(String curso) throws IOException {
        ExcelReader excelReader = new ExcelReader();

        if (radiobotonCodigo.isSelected()) {
            // Limpiar nombres para evitar conflictos
            txtAreaNombreCurso.clear();
            checkNombres.getItems().clear();

            // Cargar códigos de cursos
            List<String> codigosCursos = excelReader.obtenerCodigosDeCursos(); // Método implementado en ExcelReader
            checkCodigos.getItems().clear(); // Limpia ComboBox
            checkCodigos.getItems().addAll(codigosCursos); // Agrega los nuevos datos
        } else if (radiobotonNombre.isSelected()) {
            // Limpiar códigos para evitar conflictos
            txtcodigodelcurso.clear();
            checkCodigos.getItems().clear();

            // Cargar nombres de cursos
            List<String> nombresCursos = excelReader.obtenerNombresDeCursos(); // Método implementado en ExcelReader
            checkNombres.getItems().clear(); // Limpia ComboBox
            checkNombres.getItems().addAll(nombresCursos); // Agrega los nuevos datos
        }
    }

    class ExcelReader {

        Calendar calendario = Calendar.getInstance();
        int año = calendario.get(Calendar.YEAR);
        int periodo = calendario.get(Calendar.MONTH) < 7 ? 1 : 2;

        private String ETIQUETAS_PATH = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_importados\\" + año + "\\" + periodo + "-" + año + "\\listado_de_etiquetas_de_cursos\\";
        int versionEtiquetas = obtenerUltimaSemana(ETIQUETAS_PATH, "listado\\_\\(Semana_\\d+\\)\\.xlsx", "Semana", "xlsx");

        private String PROG_INSTITUCIONAL_PATH = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_importados\\" + año + "\\" + periodo + "-" + año + "\\programa_institucional\\";
        int versionProg = obtenerUltimaSemana(PROG_INSTITUCIONAL_PATH, "programa_institucional\\_\\(Semana_\\d+\\)\\.xlsx", "Semana", "xlsx");

        private String NUEVO_EXCEL_PATH = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos\\" + año + "\\" + periodo + "-" + año + "\\";
        int versioCondensado = obtenerUltimaSemana(NUEVO_EXCEL_PATH, "condensado\\_\\(version_\\d+\\)\\.xlsx", "version", "xlsx");

        public ExcelReader() {
            ETIQUETAS_PATH += "listado_(Semana_" + versionEtiquetas + ").xlsx";
            PROG_INSTITUCIONAL_PATH += "programa_institucional_(Semana_" + versionProg + ").xlsx";
            NUEVO_EXCEL_PATH += "condensado_(version_" + versioCondensado + ").xlsx";
        }

        public List<String> obtenerCodigosDeCursos() throws IOException {
            List<String> codigos = new ArrayList<>();
            try (FileInputStream fis = new FileInputStream(ETIQUETAS_PATH); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0); // Asume que está en la primera hoja

                for (Row row : sheet) {
                    if (row.getRowNum() <= 2) {
                        continue; // Omite encabezados
                    }
                    org.apache.poi.ss.usermodel.Cell codigoCell = row.getCell(0); // Columna con código del curso
                    if (codigoCell != null && !codigoCell.getStringCellValue().trim().isEmpty()) {
                        codigos.add(codigoCell.getStringCellValue().trim());
                    }
                }
            }
            return codigos;
        }

        public String obtenerCodigoPorNombre(String nombreCurso) throws IOException {
            try (FileInputStream fis = new FileInputStream(ETIQUETAS_PATH); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0); // Asume que está en la primera hoja

                for (Row row : sheet) {
                    if (row.getRowNum() <= 0) {
                        continue; // Omite encabezados
                    }

                    // Asume que el nombre está en la columna 1 y el código en la columna 0
                    org.apache.poi.ss.usermodel.Cell nombreCell = row.getCell(1); // Columna con nombre del curso
                    org.apache.poi.ss.usermodel.Cell codigoCell = row.getCell(0); // Columna con código del curso

                    if (nombreCell != null && codigoCell != null) {
                        if (nombreCell.getStringCellValue().equalsIgnoreCase(nombreCurso)) {
                            return codigoCell.getStringCellValue(); // Devuelve el código cuando encuentra el nombre
                        }
                    }
                }
            }
            return null; // Si no se encuentra el curso
        }

        public List<String> obtenerNombresDeCursos() throws IOException {
            List<String> nombres = new ArrayList<>();
            try (FileInputStream fis = new FileInputStream(ETIQUETAS_PATH); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0); // Asume que está en la primera hoja

                for (Row row : sheet) {
                    if (row.getRowNum() <= 2) {
                        continue; // Omite encabezados
                    }
                    org.apache.poi.ss.usermodel.Cell nombreCell = row.getCell(1); // Columna con nombre del curso
                    if (nombreCell != null && !nombreCell.getStringCellValue().trim().isEmpty()) {
                        nombres.add(nombreCell.getStringCellValue().trim());
                    }
                }
            }
            return nombres;
        }

        // Método para obtener nombres completos del nuevo Excel
        public List<String> obtenerDocentesAcreditadosPorCurso(String nombreCurso) throws IOException {
            List<String> nombresAcreditados = new ArrayList<>();

            try (FileInputStream fis = new FileInputStream(NUEVO_EXCEL_PATH); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0); // Asume que está en la primera hoja

                for (Row row : sheet) {
                    // Omite las filas de encabezado si es necesario
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    // Leer las columnas necesarias
                    org.apache.poi.ss.usermodel.Cell nombreCell = row.getCell(2); // Columna "nombre"
                    org.apache.poi.ss.usermodel.Cell apellido1Cell = row.getCell(0); // Columna "apellido1"
                    org.apache.poi.ss.usermodel.Cell apellido2Cell = row.getCell(1); // Columna "apellido2"
                    org.apache.poi.ss.usermodel.Cell cursoCell = row.getCell(7); // Columna "Nombre del Curso/Evento"
                    org.apache.poi.ss.usermodel.Cell acreditadoCell = row.getCell(11); // Columna "Acreditó"

                    // Validar datos
                    if (cursoCell != null && acreditadoCell != null && nombreCell != null
                            && apellido1Cell != null && apellido2Cell != null) {

                        String curso = cursoCell.getStringCellValue().split("\\.")[1].trim();
                        String acreditado = acreditadoCell.getStringCellValue();
                        String nombre = nombreCell.getStringCellValue();
                        String nombreCompleto = nombreCell.getStringCellValue() + " " + apellido1Cell.getStringCellValue() + " " + apellido2Cell.getStringCellValue();

                        // Filtrar por curso y acreditación
                        if (curso.equalsIgnoreCase(nombreCurso) && acreditado.equalsIgnoreCase("Si")) {
                            nombresAcreditados.add(nombreCompleto);

                        }
                    }
                }
            }

            return nombresAcreditados;
        }

        // Método para buscar Nombre del curso en el archivo "Etiquetas"
        public String buscarCurso(String codigoCurso) throws IOException {

            try (FileInputStream fis = new FileInputStream(ETIQUETAS_PATH); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0); // Asume que está en la primera hoja

                for (Row row : sheet) {
                    org.apache.poi.ss.usermodel.Cell codigoCell = row.getCell(0); // Columna del código del curso
                    org.apache.poi.ss.usermodel.Cell NombreCursoCell = row.getCell(1); // Columna del Nombre

                    if (codigoCell != null && codigoCell.getStringCellValue().equals(codigoCurso)) {
                        return NombreCursoCell != null ? NombreCursoCell.getStringCellValue() : "";
                    }
                }
            }
            return null; // No se encontró el código
        }

        // Método para buscar detalles del curso en el archivo "Prog-Institucional"
        public Map<String, String> buscarDetallesCurso(String nombreCurso) throws IOException {
            Map<String, String> datosCurso = new HashMap<>();

            try (FileInputStream file = new FileInputStream(PROG_INSTITUCIONAL_PATH)) {
                Workbook workbook = new XSSFWorkbook(file);
                Sheet sheet = workbook.getSheetAt(0);

                for (Row row : sheet) {
                    // Omite las filas antes de la fila 9 (índice 8 en términos de programación)
                    if (row.getRowNum() < 8) {
                        continue;
                    }

                    org.apache.poi.ss.usermodel.Cell nombreCursoCell = row.getCell(1); // Columna "Nombre de los evento"
                    //System.out.println("596: " +nombreCursoCell.getStringCellValue());
                    if (nombreCursoCell != null && nombreCursoCell.getCellType() == CELL_TYPE_STRING
                            && nombreCursoCell.getStringCellValue().equalsIgnoreCase(nombreCurso)) {

                        // Obtenemos cada celda relevante de la fila, usando la ruta completa de `Cell`
                        org.apache.poi.ss.usermodel.Cell competenciasCell = row.getCell(3); // Columna "Competencias a desarrollar"
                        org.apache.poi.ss.usermodel.Cell fechaCursoCell = row.getCell(4);    // Columna "Periodo de Realización"
                        org.apache.poi.ss.usermodel.Cell horasCursoCell = row.getCell(6);    // Columna "No. de horas x Curso"
                        org.apache.poi.ss.usermodel.Cell nombreInstructorCell = row.getCell(7); // Columna "Facilitador(a)"

                        // Procesamos cada celda según su tipo y añadimos validaciones adicionales
                        datosCurso.put("competencias", (competenciasCell != null && competenciasCell.getCellType() == CELL_TYPE_STRING)
                                ? competenciasCell.getStringCellValue() : "");

                        datosCurso.put("fechaCurso", (fechaCursoCell != null && fechaCursoCell.getCellType() == CELL_TYPE_STRING)
                                ? fechaCursoCell.getStringCellValue() : "");

                        datosCurso.put("horasCurso", (horasCursoCell != null && horasCursoCell.getCellType() == CELL_TYPE_NUMERIC)
                                ? String.valueOf((int) horasCursoCell.getNumericCellValue()) : "");

                        datosCurso.put("nombreInstructor", (nombreInstructorCell != null && nombreInstructorCell.getCellType() == CELL_TYPE_STRING)
                                ? nombreInstructorCell.getStringCellValue() : "");

                        break; // Detener el bucle después de encontrar el curso
                    }
                }
                workbook.close();
            }

            return datosCurso.isEmpty() ? null : datosCurso;
        }

        // Método auxiliar para obtener el valor de una celda como cadena de texto
        private String getCellStringValue(org.apache.poi.ss.usermodel.Cell cell) {
            if (cell == null) {
                return "";
            }
            switch (cell.getCellType()) { // Usa getCellType() en POI 3.12
                case CELL_TYPE_STRING:
                    return cell.getStringCellValue();
                case CELL_TYPE_NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case CELL_TYPE_BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                default:
                    return ""; // Para otros tipos de celdas o celdas de error
            }
        }

        public boolean guardarDetallesCurso(String nombreCurso, Map<String, String> datosActualizados) throws IOException {
            boolean actualizado = false;

            try (FileInputStream file = new FileInputStream(PROG_INSTITUCIONAL_PATH); Workbook workbook = new XSSFWorkbook(file)) {
                Sheet sheet = workbook.getSheetAt(0);

                for (Row row : sheet) {
                    // Omite las filas antes de la fila 9 (índice 8)
                    if (row.getRowNum() < 8) {
                        continue;
                    }

                    // Ubica la fila que corresponde al curso especificado
                    org.apache.poi.ss.usermodel.Cell nombreCursoCell = row.getCell(1);
                    if (nombreCursoCell != null && nombreCursoCell.getCellType() == CELL_TYPE_STRING
                            && nombreCursoCell.getStringCellValue().equalsIgnoreCase(nombreCurso)) {

                        // Actualizar los datos en las celdas correspondientes
                        if (datosActualizados.containsKey("competencias")) {
                            org.apache.poi.ss.usermodel.Cell competenciasCell = row.getCell(3);
                            if (competenciasCell == null) {
                                competenciasCell = row.createCell(3);
                            }
                            competenciasCell.setCellValue(datosActualizados.get("competencias"));
                        }

                        if (datosActualizados.containsKey("fechaCurso")) {
                            org.apache.poi.ss.usermodel.Cell fechaCursoCell = row.getCell(4);
                            if (fechaCursoCell == null) {
                                fechaCursoCell = row.createCell(4);
                            }
                            fechaCursoCell.setCellValue(datosActualizados.get("fechaCurso"));
                        }

                        if (datosActualizados.containsKey("horasCurso")) {
                            org.apache.poi.ss.usermodel.Cell horasCursoCell = row.getCell(6);
                            if (horasCursoCell == null) {
                                horasCursoCell = row.createCell(6);
                            }
                            horasCursoCell.setCellValue(Integer.parseInt(datosActualizados.get("horasCurso")));
                        }

                        if (datosActualizados.containsKey("nombreInstructor")) {
                            org.apache.poi.ss.usermodel.Cell nombreInstructorCell = row.getCell(7);
                            if (nombreInstructorCell == null) {
                                nombreInstructorCell = row.createCell(7);
                            }
                            nombreInstructorCell.setCellValue(datosActualizados.get("nombreInstructor"));
                        }

                        actualizado = true;
                        break;
                    }
                }

                if (actualizado) {
                    // Guardar los cambios en el archivo
                    try (FileOutputStream fos = new FileOutputStream(PROG_INSTITUCIONAL_PATH)) {
                        workbook.write(fos);
                    }
                }
            }

            return actualizado;
        }

    }
}
