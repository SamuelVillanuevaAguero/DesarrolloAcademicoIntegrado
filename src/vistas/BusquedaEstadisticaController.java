package vistas;

import static com.aspose.cells.PropertyType.BOOLEAN;
import static com.aspose.cells.PropertyType.NUMBER;
import static com.aspose.cells.PropertyType.STRING;
import com.aspose.cells.SaveFormat;
import static com.sun.javafx.scene.control.skin.FXVK.Type.NUMERIC;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Year;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.MouseEvent;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utilerias.busqueda.filaDato;
import utilerias.general.ControladorGeneral;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.DateTimeParseException;
import java.time.format.ResolverStyle;
import java.time.temporal.ChronoField;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.traversal.NodeFilter;

/*import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPCell;*/
public class BusquedaEstadisticaController implements Initializable {

    @FXML
    private Button botonCerrar;
    @FXML
    private Button botonMinimizar;
    @FXML
    private Button botonRegresar;
    @FXML
    private Label botonActualizar;
    @FXML
    private Button botonLimpiar;
    @FXML
    private Button botonBuscar;
    @FXML
    private Button botonExportar;
    @FXML
    private Label botonActualizarDatos;

    @FXML
    private ComboBox<String> comboTipoCapacitacion;
    @FXML
    private ComboBox<String> comboDepartamento;
    @FXML
    private ComboBox<String> comboAcreditacion;
    @FXML
    private ComboBox<String> comboNivel;
    @FXML
    private ComboBox<Integer> comboAño;
    @FXML
    private ComboBox<String> comboPeriodo;
    @FXML
    private ComboBox<String> comboFormato;

    // TABLA ------------------------------------------------------------------------
    @FXML
    private TableView<filaDato> tabla;

    @FXML
    private TableColumn<filaDato, Integer> columnaAño;
    @FXML
    private TableColumn<filaDato, String> columnaPeriodo;
    @FXML
    private TableColumn<filaDato, String> columnaNombre;
    @FXML
    private TableColumn<filaDato, String> columnaApellidoPaterno;
    @FXML
    private TableColumn<filaDato, String> columnaApellidoMaterno;
    @FXML
    private TableColumn<filaDato, String> columnaDepartamentoLicenciatura;
    @FXML
    private TableColumn<filaDato, String> columnaDepartamentoPosgrado;
    @FXML
    private TableColumn<filaDato, String> columnaAcreditado;
    @FXML
    private TableColumn<filaDato, String> columnaTipoCapacitacion;
    @FXML
    private TableColumn<filaDato, Integer> columnaNumeroCursos;

    @FXML
    private Label numeroTotalDocentes;
    @FXML
    private Label docentesTomandoCursos;
    @FXML
    private Label porcentajeDocentesCapacitados;
    @FXML
    private Label NumeroDocentesNivel;

    private ObservableList data;

    public void exportarExcelAPDF(String rutaArchivo, File fileToSave) {
        try {
            // Cargar el archivo de Excel
            com.aspose.cells.Workbook workbook = new com.aspose.cells.Workbook(rutaArchivo);

            // Guardar el archivo como PDF
            workbook.save(fileToSave.getAbsolutePath(), SaveFormat.PDF);
            System.out.println("Archivo exportado exitosamente como PDF en: " + fileToSave.getAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static Cell obtenerCeldaDebajo(String filePath) {
        try (FileInputStream fileInputStream = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Obtener la primera hoja del libro
            Sheet sheet = workbook.getSheetAt(0);

            // Buscar la celda con el texto "NOMBRE DEL DOCENTE"
            for (Row row : sheet) {
                for (Cell cell : row) {
                    // Usar CellType en lugar de constantes deprecated
                    if (cell.getCellType() == Cell.CELL_TYPE_STRING && "NOMBRE DEL DOCENTE".equals(cell.getStringCellValue())) {

                        // Obtener el índice de la fila justo debajo
                        int rowIndexDebajo = row.getRowNum() + 1;
                        Row rowDebajo = sheet.getRow(rowIndexDebajo);

                        // Crear la fila si no existe
                        if (rowDebajo == null) {
                            rowDebajo = sheet.createRow(rowIndexDebajo);
                        }

                        // Obtener la celda en la misma columna
                        int columnIndex = cell.getColumnIndex();
                        Cell cellDebajo = rowDebajo.getCell(columnIndex);

                        // Crear la celda si no existe
                        if (cellDebajo == null) {
                            cellDebajo = rowDebajo.createCell(columnIndex);
                        }

                        // Actualizar la celda debajo con el valor "Samito"
                        cellDebajo.setCellValue("Samito");
                        System.out.println("Escrito en la celda debajo de 'NOMBRE DEL DOCENTE'");

                        // Retornar la celda justo debajo
                        return cellDebajo;
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return null; // Si no se encuentra la celda "NOMBRE DEL DOCENTE"
    }

    public void escribirExcel(String rutaArchivo, Stage stage) {
        try (FileInputStream fileInputStream = new FileInputStream(rutaArchivo); Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Obtener la primera hoja

            // Estilo para datos normales  
            CellStyle estiloNormal = workbook.createCellStyle();
            estiloNormal.setAlignment(CellStyle.ALIGN_CENTER);
            estiloNormal.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            estiloNormal.setBorderTop(CellStyle.BORDER_THIN);
            estiloNormal.setBorderBottom(CellStyle.BORDER_THIN);
            estiloNormal.setBorderLeft(CellStyle.BORDER_THIN);
            estiloNormal.setBorderRight(CellStyle.BORDER_THIN);

            // Estilo para encabezados (en negrita)  
            CellStyle estiloEncabezado = workbook.createCellStyle();
            estiloEncabezado.cloneStyleFrom(estiloNormal);
            Font fuenteNegrita = workbook.createFont();

            fuenteNegrita.setBold(true);
            estiloEncabezado.setFont(fuenteNegrita);

            // Crear encabezados si no existen
            String[] encabezados = {"NO.", "NOMBRE DEL DOCENTE", "NO. CURSO", "F.DOCENTE", "A. DOCENTE"};
            if (sheet.getRow(9) == null) {
                Row headerRow = sheet.createRow(9);
                for (int i = 0; i < encabezados.length; i++) {
                    Cell cell = headerRow.createCell(i + 1);
                    cell.setCellValue(encabezados[i]);
                    cell.setCellStyle(estiloEncabezado);
                }
            }

            // Encontrar la siguiente fila vacía para agregar datos
            int rowNum = 10;
            while (sheet.getRow(rowNum) != null) {
                rowNum++;
            }

            // Llenar datos  
            Row row = sheet.createRow(rowNum);
            String[] datos = {"1", "Samuelin", "45", "Si", "No"};
            for (int i = 0; i < datos.length; i++) {
                Cell cell = row.createCell(i + 1);
                cell.setCellValue(datos[i]);
                cell.setCellStyle(estiloNormal);
            }

            // Ajustar ancho de columnas  
            for (int i = 0; i < encabezados.length; i++) {
                sheet.autoSizeColumn(i + 1);
            }

            // Usar FileChooser para seleccionar la ubicación de exportación
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Guardar como");
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
            fileChooser.setInitialFileName("copia_" + new File(rutaArchivo).getName()); // Nombre por defecto

            File fileToSave = fileChooser.showSaveDialog(stage);
            if (fileToSave != null) {
                try (FileOutputStream outputStream = new FileOutputStream(fileToSave)) {
                    workbook.write(outputStream);
                }
                System.out.println("Datos añadidos exitosamente al archivo Excel en: " + fileToSave.getAbsolutePath());
            } else {
                System.out.println("Exportación cancelada.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void exportarArchivo(String rutaArchivo, Stage stage, int formato) {
        try (FileInputStream fileInputStream = new FileInputStream(rutaArchivo); Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Obtener la primera hoja

            // Usar FileChooser para seleccionar la ubicación de exportación
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Guardar como");

            // Definir el tipo de archivo según el parámetro formato
            if (formato == 1) {
                // Formato PDF
                fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("PDF Files", "*.pdf"));
                fileChooser.setInitialFileName("copia_" + new File(rutaArchivo).getName().replace(".xlsx", ".pdf"));
            } else if (formato == 2) {
                // Formato Excel
                fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
                fileChooser.setInitialFileName("copia_" + new File(rutaArchivo).getName());
            }

            File fileToSave = fileChooser.showSaveDialog(stage);
            if (fileToSave != null) {
                if (formato == 1) {
                    // Guardar como PDF usando Aspose.Cells
                    exportarAPDF(rutaArchivo, fileToSave);
                    System.out.println("Archivo exportado exitosamente como PDF en: " + fileToSave.getAbsolutePath());
                } else if (formato == 2) {
                    // Guardar como Excel
                    try (FileOutputStream outputStream = new FileOutputStream(fileToSave)) {
                        workbook.write(outputStream);
                    }
                    System.out.println("Archivo exportado exitosamente como Excel en: " + fileToSave.getAbsolutePath());
                }
            } else {
                System.out.println("Exportación cancelada.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void exportarAPDF(String rutaArchivo, File fileToSave) {
        try {
            // Cargar el archivo Excel con Aspose.Cells
            com.aspose.cells.Workbook workbook = new com.aspose.cells.Workbook(rutaArchivo);

            // Guardar el archivo Excel como PDF
            workbook.save(fileToSave.getAbsolutePath(), SaveFormat.PDF);

            System.out.println("Archivo exportado exitosamente como PDF en: " + fileToSave.getAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public int contarFilasUnicas(ObservableList<filaDato> data) {
        // Utilizamos un Set para almacenar combinaciones únicas
        Set<String> setUnico = new HashSet<>();

        // Recorremos los datos y agregamos la combinación al Set
        for (filaDato fila : data) {
            String key = fila.getNombre() + " " + fila.getApellidoPaterno() + " " + fila.getApellidoMaterno();
            setUnico.add(key);  // Set no permite duplicados
        }

        // El tamaño del Set será el número de filas únicas
        return setUnico.size();
    }

    /*private int obtenerTotalDocentes(String ruta, int año, int periodo) {
        try {
            FileInputStream file = new FileInputStream(ruta);
            Workbook libro = new XSSFWorkbook(file);

            Sheet hoja = libro.getSheetAt(0);

            return (int) Double.parseDouble(hoja.getRow(3).getCell(2).toString());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
        }
        return 0;
    }*/
    private int obtenerTotalDocentes(String ruta, int año, int periodo) {
        try {
            // Abrir el archivo Excel
            FileInputStream file = new FileInputStream(ruta);
            Workbook libro = new XSSFWorkbook(file);

            // Buscar la hoja por nombre basado en el año
            Sheet hoja = libro.getSheet(String.valueOf(año));
            if (hoja == null) {
                System.err.println("No se encontró una hoja con el nombre: " + año);
                return 0; // Salir si no se encuentra la hoja
            }

            // Leer el valor de la celda en la fila 3, columna 2 (Indexada desde 0)
            return (int) Double.parseDouble(hoja.getRow(periodo + 2).getCell(1).toString());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Archivo no encontrado: " + ruta, ex);
        } catch (IOException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error al leer el archivo Excel", ex);
        } catch (NullPointerException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error: Celda o fila no encontrada", ex);
        } catch (NumberFormatException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error al convertir el valor de la celda a número", ex);
        }
        return 0; // Devolver 0 si ocurre algún error
    }

    private ObservableList<filaDato> readExcelData(String filePath) throws IOException {
        ObservableList<filaDato> data = FXCollections.observableArrayList();
        try (FileInputStream file = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            // Second pass: Create filaDato objects
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                System.out.println(sheet.getLastRowNum());
                Row row = sheet.getRow(i);
                if (row != null) {
                    System.out.println(row);
                    int año = 0;
                    String acreditado = "";
                    int noCursos = 0;

                    String nombre = row.getCell(0).getStringCellValue();
                    String apellidoPaterno = row.getCell(1).getStringCellValue();
                    String apellidoMaterno = row.getCell(2).getStringCellValue();
                    String periodo = row.getCell(6).getStringCellValue();
                    String departamento = row.getCell(8).getStringCellValue();
                    String posgrado = row.getCell(9).getStringCellValue();
                    String tipoCapacitacion = row.getCell(13).getStringCellValue();
                    // Assuming this cell has accreditation info
                    System.out.println(i);
                    // Extract year from the period
                    String[] dates = periodo.split("-");
                    if (dates[0].trim().split("/").length > 1) {
                        año = Integer.parseInt(dates[0].trim().split("/")[2]);
                        acreditado = row.getCell(12).getStringCellValue();
                    }

                    // Get number of courses for the instructor
                    // Create and add filaDato object
                    data.add(new filaDato(año, periodo, nombre, apellidoPaterno, apellidoMaterno, departamento, posgrado, acreditado, tipoCapacitacion, noCursos));
                }
            }
        }
        return data;
    }

    /*private ObservableList<filaDato> readAllExcelFiles(Integer año, Integer periodo) {
        ObservableList<filaDato> data = FXCollections.observableArrayList();
        Map<String, filaDato> dataMap = new HashMap<>();

        // Directorio base  
        String baseDir = ControladorGeneral.obtenerRutaDeEjecusion()+"\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos";
        Path searchPath;

        // Construir la ruta del directorio basado en año y periodo  
        if (año != 0) {
            if (periodo == 0) {
                searchPath = Paths.get(baseDir, año.toString());
            } else {
                searchPath = Paths.get(baseDir, año.toString(), periodo.toString()+"-"+año);
            }
        } else {
            searchPath = Paths.get(baseDir);
        }

        try {
            Stream<Path> pathStream = Files.walk(searchPath)
                    .filter(path -> path.toString().endsWith("condensado_(version_"+numeroVersion+").xlsx"));

            pathStream.forEach(path -> {
                System.out.println("Hoja n");
                try (FileInputStream file = new FileInputStream(path.toFile()); Workbook workbook = WorkbookFactory.create(file)) {

                    Sheet sheet = workbook.getSheetAt(0);

                    // Recorrer las filas del archivo Excel
                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);

                        if (row == null) {
                            continue;
                        }

                        int currentAño = 0;
                        int currentPeriodo = 0;

                        // Extracción de periodo y año  
                        String periodoCell = row.getCell(6).getStringCellValue();
                        String[] dates = periodoCell.split("-");
                        if (dates[0].trim().split("/").length > 1) {
                            currentAño = Integer.parseInt(dates[0].trim().split("/")[2]);
                            currentPeriodo = Integer.parseInt(dates[0].trim().split("/")[1]) <= 7 ? 1 : 2;
                        }

                        boolean passAñoFilter = currentAño == año || año == 0;
                        boolean passPeriodoFilter = periodo == 0 || currentPeriodo == periodo;

                        if (passAñoFilter && passPeriodoFilter) {
                            String nombre = row.getCell(0).getStringCellValue();
                            String apellidoPaterno = row.getCell(1).getStringCellValue();
                            String apellidoMaterno = row.getCell(2).getStringCellValue();
                            String departamento = row.getCell(8).getStringCellValue();
                            String posgrado = row.getCell(9).getStringCellValue();
                            String acreditado = row.getCell(12).getStringCellValue();
                            String tipoCapacitacion = row.getCell(13).getStringCellValue();
                            String key = currentAño + " " + periodoCell + " " + nombre + " " + apellidoPaterno + " " + apellidoMaterno;

                            // Verificar si ya existe el registro en el Map
                            if (dataMap.containsKey(key)) {
                                // Incrementar el contador de cursos si ya existe
                                filaDato existingDato = dataMap.get(key);
                                existingDato.setNoCursos(existingDato.getNoCursos() + 1);
                            } else {
                                // Si no existe, crear un nuevo objeto filaDato con noCursos = 1
                                filaDato newDato = new filaDato(currentAño, periodoCell, nombre,
                                        apellidoPaterno, apellidoMaterno, departamento, posgrado,
                                        acreditado, tipoCapacitacion, 1);
                                dataMap.put(key, newDato);
                            }
                        }
                    }
                } catch (Exception e) {
                    Logger.getLogger(BusquedaEstadisticaController.class.getName())
                            .log(Level.SEVERE, "Error reading file: " + path, e);
                }
            });
        } catch (IOException e) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName())
                    .log(Level.SEVERE, "Error walking directory", e);
        }

        // Convertir el Map a una lista observable
        data.addAll(dataMap.values());
        return data;
    }*/

    /*private ObservableList<filaDato> readAllExcelFiles(Integer año, Integer periodo, String tipoCapacitacion, String departamento, String acreditacion, String nivel) {
        ObservableList<filaDato> data = FXCollections.observableArrayList();
        Map<String, filaDato> dataMap = new HashMap<>();

        // Directorio base  
        String baseDir = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos";
        Path searchPath;

        // Construir la ruta del directorio basado en año y periodo  
        if (año != 0) {
            searchPath = periodo == 0 ? Paths.get(baseDir, año.toString()) : Paths.get(baseDir, año.toString(), periodo.toString());
        } else {
            searchPath = Paths.get(baseDir);
        }

        try {
            // Filtrar archivos Excel por nombre
            Stream<Path> pathStream = Files.walk(searchPath)
                    .filter(path -> path.toString().endsWith("Acreditacion_Docente.xlsx"));

            pathStream.forEach(path -> {
                try (FileInputStream file = new FileInputStream(path.toFile()); Workbook workbook = WorkbookFactory.create(file)) {

                    Sheet sheet = workbook.getSheetAt(0);

                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);
                        if (row == null) {
                            continue;
                        }

                        int currentAño = 0;
                        int currentPeriodo = 0;

                        // Extracción de año y periodo
                        String periodoCell = row.getCell(6).getStringCellValue();
                        String[] dates = periodoCell.split("-");
                        if (dates[0].trim().split("/").length > 1) {
                            currentAño = Integer.parseInt(dates[0].trim().split("/")[2]);
                            currentPeriodo = Integer.parseInt(dates[0].trim().split("/")[1]) <= 7 ? 1 : 2;
                        }

                        boolean passAñoFilter = currentAño == año || año == 0;
                        boolean passPeriodoFilter = periodo == 0 || currentPeriodo == periodo;

                        // Aplicar filtros de ComboBox
                        String filaTipoCapacitacion = row.getCell(13).getStringCellValue();
                        String filaDepartamento = row.getCell(8).getStringCellValue();
                        String filaPosgrado = row.getCell(9).getStringCellValue();
                        String filaAcreditacion = row.getCell(12).getStringCellValue();

                        boolean passTipoCapacitacion = tipoCapacitacion == null || tipoCapacitacion.equals(filaTipoCapacitacion);
                        boolean passDepartamento = departamento == null || departamento.equals(filaDepartamento);
                        boolean passAcreditacion = acreditacion == null || acreditacion.equalsIgnoreCase(filaAcreditacion) || acreditacion.equals("Ambos");
                        boolean passNivel = nivel == null || (nivel.equals("Licenciatura") && filaPosgrado.equals("No")) || (nivel.equals("Posgrado") && filaPosgrado.equals("Sí"));

                        // Aplicar todos los filtros
                        if (passAñoFilter && passPeriodoFilter && passTipoCapacitacion && passDepartamento && passAcreditacion && passNivel) {
                            String nombre = row.getCell(0).getStringCellValue();
                            String apellidoPaterno = row.getCell(1).getStringCellValue();
                            String apellidoMaterno = row.getCell(2).getStringCellValue();

                            // Crear clave única para el Map
                            String key = currentAño + " " + periodoCell + " " + nombre + " " + apellidoPaterno + " " + apellidoMaterno + " " + filaDepartamento + " " + filaPosgrado + " " + filaAcreditacion + " " + filaTipoCapacitacion;

                            // Verificar si ya existe el registro en el Map
                            if (dataMap.containsKey(key)) {
                                filaDato existingDato = dataMap.get(key);
                                existingDato.setNoCursos(existingDato.getNoCursos() + 1);
                            } else {
                                filaDato newDato = new filaDato(currentAño, periodoCell, nombre,
                                        apellidoPaterno, apellidoMaterno, filaDepartamento, filaPosgrado,
                                        filaAcreditacion, filaTipoCapacitacion, 1);
                                dataMap.put(key, newDato);
                            }
                        }
                    }
                } catch (Exception e) {
                    Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error leyendo archivo: " + path, e);
                }
            });
        } catch (IOException e) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error recorriendo directorio", e);
        }

        data.addAll(dataMap.values());
        return data;
    }*/
    private ObservableList<filaDato> readAllExcelFiles(Integer año, Integer periodo, String tipoCapacitacion, String departamento, String acreditacion, String nivel) throws InvalidFormatException {
        ObservableList<filaDato> data = FXCollections.observableArrayList();
        Map<String, filaDato> dataMap = new HashMap<>();
        Map<String, filaDato> dataMap2 = new HashMap<>();

        // Directorio base
        String baseDir = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos";
        Path searchPath;

        // Construir la ruta del directorio basado en año y periodo
        if (año != 0) {
            searchPath = periodo == 0 ? Paths.get(baseDir, año.toString()) : Paths.get(baseDir, año.toString(), periodo.toString() + "-" + año);
        } else {
            searchPath = Paths.get(baseDir);
        }

        try {
            // Determinar el número de la versión más alta
            String carpetaDestino = searchPath.toString();
            int maxVersion = determinarNumeroDeVersion(carpetaDestino);

            // Generar el nombre del archivo con la versión más alta
            String archivoConMayorVersion = carpetaDestino + "\\condensado_(version_" + maxVersion + ").xlsx";
            System.out.println(archivoConMayorVersion);

            File archivo = new File(archivoConMayorVersion);
            if (!archivo.exists()) {
                Logger.getLogger(BusquedaEstadisticaController.class.getName())
                        .log(Level.WARNING, "No se encontró el archivo con la versión más alta: " + archivoConMayorVersion);
                return data; // Retornar vacío si no se encuentra el archivo
            }

            // Procesar el archivo Excel
            try (FileInputStream file = new FileInputStream(archivo); Workbook workbook = WorkbookFactory.create(file)) {
                Sheet sheet = workbook.getSheetAt(0); // Usamos la primera hoja del libro Excel

                int contadorPosgrado = 0;
                for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Iterar desde la segunda fila

                    Row row = sheet.getRow(i);
                    if (row == null) {
                        continue;
                    }

                    int currentAño = 0;
                    int currentPeriodo = 0;
                    String periodoCell = "";

                    try {
                        // Leer la celda de la fecha y parsear al formato "dd/MMM/yyyy" con meses en letras
                        String[] mes = row.getCell(10).getStringCellValue().split(" ");
                        periodoCell = mes[mes.length - 1].substring(0, 3).toLowerCase();


                        if (!periodoCell.isEmpty()) {
                            // Crear un formateador que soporte abreviaturas de meses en español
                            //DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MMM/yyyy", Locale.forLanguageTag("es"));

                            // Parsear la fecha con el formateador
                            //LocalDate date = LocalDate.parse(periodoCell, formatter);
                            currentAño = año;
                            currentPeriodo = extraerMes(periodoCell) < 7 ? 1 : 2; // Determinar el período
                            periodoCell = currentPeriodo == 1 ? "ENERO - JULIO" : "AGOSTO - DICIEMBRE";
                        }
                    } catch (DateTimeParseException e) {
                        Logger.getLogger(BusquedaEstadisticaController.class.getName())
                                .log(Level.WARNING, "Error al parsear la fecha en la fila " + i + " del archivo: " + archivoConMayorVersion, e);
                        continue; // Saltar esta fila si la fecha no es válida
                    }

                    // Validar filtros de año y período
                    boolean passAñoFilter = currentAño == año || año == 0;
                    boolean passPeriodoFilter = periodo == 0 || currentPeriodo == periodo;

                    // Leer las celdas con los valores necesarios
                    String filaTipoCapacitacion = getCellValueAsString(row.getCell(8));
                    String filaDepartamento = getCellValueAsString(row.getCell(5));
                    String filaPosgrado = getCellValueAsString(row.getCell(5)).equalsIgnoreCase("POSGRADO") ? "si" : "no";
                    String filaAcreditacion = getCellValueAsString(row.getCell(11));

                    // Validar filtros de ComboBox
                    boolean passTipoCapacitacion = tipoCapacitacion == null || tipoCapacitacion.equalsIgnoreCase(filaTipoCapacitacion);
                    boolean passDepartamento = departamento == null || departamento.equalsIgnoreCase(filaDepartamento);
                    boolean passAcreditacion = acreditacion == null || acreditacion.equalsIgnoreCase(filaAcreditacion) || acreditacion.equalsIgnoreCase("Ambos");
                    boolean passNivel = nivel == null
                            || (nivel.equals("Licenciatura") && filaPosgrado.equalsIgnoreCase("No"))
                            || (nivel.equals("Posgrado") && filaPosgrado.equalsIgnoreCase("Si"));

                    // Aplicar todos los filtros
                    if (passAñoFilter && passPeriodoFilter && passTipoCapacitacion && passDepartamento && passAcreditacion && passNivel) {
                        String nombre = getCellValueAsString(row.getCell(2));
                        String apellidoPaterno = getCellValueAsString(row.getCell(0));
                        String apellidoMaterno = getCellValueAsString(row.getCell(1));

                        // Crear clave única para el Map
                        String key = currentAño + " " + periodoCell + " " + nombre + " " + apellidoPaterno + " " + apellidoMaterno + " " + filaDepartamento + " " + filaPosgrado + " " + filaAcreditacion + " " + filaTipoCapacitacion;
                        String key2 = nombre + " " + apellidoPaterno + " " + apellidoMaterno;

                        if (!dataMap2.containsKey(key2) && filaPosgrado.equalsIgnoreCase("si")) {
                            contadorPosgrado += 1;
                            dataMap2.put(key2, new filaDato(0, periodoCell, nombre, apellidoPaterno, apellidoMaterno, departamento, departamento, acreditacion, tipoCapacitacion, 0));
                        }

                        // Verificar si ya existe el registro en el Map
                        if (dataMap.containsKey(key)) {
                            filaDato existingDato = dataMap.get(key);
                            existingDato.setNoCursos(existingDato.getNoCursos() + 1);
                        } else {
                            filaDato newDato = new filaDato(currentAño, periodoCell, nombre,
                                    apellidoPaterno, apellidoMaterno, filaDepartamento, filaPosgrado,
                                    filaAcreditacion, filaTipoCapacitacion, 1);
                            dataMap.put(key, newDato);
                        }
                    }
                }

                NumeroDocentesNivel.setText(String.valueOf(contadorPosgrado));
            }

        } catch (IOException e) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, "Error recorriendo directorio", e);
        }

        data.addAll(dataMap.values());
        return data;
    }

    public int extraerMes(String fecha) {
        // Crear un mapa para convertir los nombres de los meses en números  
        Map<String, Integer> meses = new HashMap<>();
        meses.put("ene", 1);
        meses.put("feb", 2);
        meses.put("mar", 3);
        meses.put("abr", 4);
        meses.put("may", 5);
        meses.put("jun", 6);
        meses.put("jul", 7);
        meses.put("ago", 8);
        meses.put("sep", 9);
        meses.put("oct", 10);
        meses.put("nov", 11);
        meses.put("dic", 12);

        return meses.getOrDefault(fecha, -1); // Retorna -1 si el mes no es válido  
    }

    public static int extraerAnio(String fecha) {
        // Separar la cadena por guiones  
        String[] partes = fecha.split("-");
        if (partes.length != 3) {
            throw new IllegalArgumentException("La fecha no tiene el formato correcto");
        }

        // El año es la tercera parte  
        try {
            return Integer.parseInt(partes[2]); // Convertir a entero  
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("El año no es válido");
        }
    }

    private int determinarNumeroDeVersion(String carpetaDestino) {
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
        File[] archivos = carpeta.listFiles((dir, name) -> name.matches("condensado_\\(version_\\d+\\)\\.xlsx"));

        if (archivos == null || archivos.length == 0) {
            return 1; // Si no hay archivos, retornamos 1 como primera versión
        }

        // Determinar la versión más alta
        int maxVersion = 0;
        Pattern pattern = Pattern.compile(".*\\(version_(\\d+)\\)\\.xlsx$"); // Patrón para extraer el número de versión

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

// Método auxiliar para obtener valores de celdas como cadenas
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        try {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING: // Para texto
                    return cell.getStringCellValue();

                case Cell.CELL_TYPE_NUMERIC: // Para números o fechas
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // Si es una fecha, formatear como dd/MMM/yyyy (nov para noviembre)
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MMM/yyyy", Locale.ENGLISH); // Usar inglés para abreviaturas estándar
                        return dateFormat.format(date);
                    } else {
                        // Si es un número, retornar como entero si no tiene decimales
                        double numericValue = cell.getNumericCellValue();
                        if (numericValue == (int) numericValue) {
                            return String.valueOf((int) numericValue);
                        }
                        return String.valueOf(numericValue);
                    }

                case Cell.CELL_TYPE_BOOLEAN: // Para booleanos
                    return String.valueOf(cell.getBooleanCellValue());

                case Cell.CELL_TYPE_FORMULA: // Para fórmulas
                    try {
                        return cell.getStringCellValue(); // Intentar leer como texto
                    } catch (IllegalStateException e) {
                        return String.valueOf(cell.getNumericCellValue()); // Si no, leer como numérico
                    }

                case Cell.CELL_TYPE_BLANK: // Si está en blanco
                    return "";

                default: // Cualquier otro tipo no reconocido
                    return "";
            }
        } catch (Exception e) {
            // Manejar errores generales y registrar advertencia
            Logger.getLogger(BusquedaEstadisticaController.class.getName())
                    .log(Level.WARNING, "Error al leer la celda: " + e.getMessage(), e);
            return "";
        }
    }

    public List<Integer> getAvailableYears(String baseDirectoryPath) {
        List<Integer> years = new ArrayList<>();
        Path basePath = Paths.get(baseDirectoryPath);

        try {
            // Listar solo las carpetas en el directorio base
            years = Files.list(basePath)
                    .filter(Files::isDirectory)
                    .map(path -> path.getFileName().toString()) // Obtener el nombre de la carpeta
                    .filter(name -> name.matches("\\d{4}")) // Filtrar solo los nombres que sean números de 4 dígitos (año)
                    .map(Integer::parseInt) // Convertir el nombre de la carpeta a entero
                    .sorted(Comparator.reverseOrder()) // Ordenar de forma descendente
                    .collect(Collectors.toList());
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("Error al listar los años en el directorio: " + baseDirectoryPath);
        }

        return years;
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {

        columnaAño = new TableColumn<>("Año");
        columnaAño.setCellValueFactory(new PropertyValueFactory<>("año"));
        columnaPeriodo = new TableColumn<>("Periodo");
        columnaPeriodo.setCellValueFactory(new PropertyValueFactory<>("periodo"));
        columnaNombre = new TableColumn<>("Nombre");
        columnaNombre.setCellValueFactory(new PropertyValueFactory<>("nombre"));
        columnaApellidoPaterno = new TableColumn<>("APP");
        columnaApellidoPaterno.setCellValueFactory(new PropertyValueFactory<>("apellidoPaterno"));
        columnaApellidoMaterno = new TableColumn<>("APM");
        columnaApellidoMaterno.setCellValueFactory(new PropertyValueFactory<>("apellidoMaterno"));
        columnaDepartamentoLicenciatura = new TableColumn<>("Licenciatura");
        columnaDepartamentoLicenciatura.setCellValueFactory(new PropertyValueFactory<>("departamentoLicenciatura"));
        columnaDepartamentoPosgrado = new TableColumn<>("Posgrado");
        columnaDepartamentoPosgrado.setCellValueFactory(new PropertyValueFactory<>("departamentoPosgrado"));
        columnaAcreditado = new TableColumn<>("Acreditado");
        columnaAcreditado.setCellValueFactory(new PropertyValueFactory<>("acreditado"));

        columnaTipoCapacitacion = new TableColumn<>("Capacitación");
        columnaTipoCapacitacion.setCellValueFactory(new PropertyValueFactory<>("tipoCapacitacion"));
        columnaNumeroCursos = new TableColumn<>("No.Cursos");
        columnaNumeroCursos.setCellValueFactory(new PropertyValueFactory<>("noCursos"));

        tabla.getColumns().addAll(columnaAño, columnaPeriodo, columnaNombre,
                columnaApellidoPaterno, columnaApellidoMaterno, columnaDepartamentoLicenciatura,
                columnaDepartamentoPosgrado, columnaAcreditado, columnaTipoCapacitacion, columnaNumeroCursos);

        try {
            Calendar calendario = Calendar.getInstance();
            int year = calendario.get(Calendar.YEAR);
            int mesActual = (calendario.get(Calendar.MONTH) + 1) < 7 ? 1 : 2;
            data = readAllExcelFiles(year, mesActual, null, null, null, null);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
        }

        tabla.setItems(data);
        int docentesCursos = contarFilasUnicas(data);

        Calendar calendario = Calendar.getInstance();
        int year = calendario.get(Calendar.YEAR);
        int mesActual = calendario.get(Calendar.MONTH) + 1;
        int totalDocentes = obtenerTotalDocentes(ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx", year, (mesActual >= 1 && mesActual <= 7) ? 1 : 2);
        double porcentajeCapacitados = (double) docentesCursos / totalDocentes * 100;

        docentesTomandoCursos.setText(String.valueOf(docentesCursos));
        numeroTotalDocentes.setText(String.valueOf(totalDocentes));

        porcentajeDocentesCapacitados.setText((int) porcentajeCapacitados + "%");
        // Inicializar opciones de ComboBoxes
        comboTipoCapacitacion.getItems().addAll("Actualización profesional", "Formación docente");
        comboDepartamento.getItems().addAll("CIENCIAS BÁSICAS", "CIENCIAS ECONÓMICO ADMINISTRATIVAS", "CIENCIAS DE LA TIERRA", "INGENIERÍA INDUSTRIAL", "METAL MECÁNICA", "QUÍMICA Y BIOQUÍMICA", "SISTEMAS COMPUTACIONALES");
        comboAcreditacion.getItems().addAll("Si", "No", "Ambos");
        comboNivel.getItems().addAll("Licenciatura", "Posgrado");

        int currentYear = Year.now().getValue();

        comboAño.getItems().addAll(getAvailableYears(ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos"));
        comboPeriodo.getItems().addAll("Enero - Julio", "Agosto - Diciembre");
        comboFormato.getItems().addAll("PDF", "EXCEL");
        
        comboAño.setValue(currentYear);
        comboPeriodo.setValue(mesActual == 2 ? "Enero - Julio" : "Agosto - Diciembre");

        botonCerrar.setOnMouseClicked(event -> {
            try {
                cerrarVentana(event);
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        });

        botonMinimizar.setOnMouseClicked(event -> minimizarVentana(event));
        botonRegresar.setOnMouseClicked(event -> {
            try {
                regresarVentana(event);
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        });

        botonActualizar.setOnMouseClicked(event -> {
            try {
                actualizarDocumentos(event);
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        });

        botonLimpiar.setOnMouseClicked(event -> {
            comboTipoCapacitacion.setValue(null);
            comboDepartamento.setValue(null);
            comboAcreditacion.setValue(null);
            comboNivel.setValue(null);
            comboAño.setValue(null);
            comboPeriodo.setValue(null);
            comboFormato.setValue(null);
            tabla.getItems().clear();
        });

        botonBuscar.setOnMouseClicked(event -> {

            System.out.println("Botón buscar...");

            String tipoCapacitacion = comboTipoCapacitacion.getValue() == null ? null : comboTipoCapacitacion.getValue().equals("Formación docente") ? "FD" : "AP";

            String departamento = comboDepartamento.getValue();
            String acreditacion = comboAcreditacion.getValue();
            String nivel = comboNivel.getValue();

            int año = comboAño.getValue() != null ? comboAño.getValue() : 0;
            int periodo = comboPeriodo.getValue() != null
                    ? comboPeriodo.getValue().equals("Enero - Julio") ? 1 : 2 : 0;

            if (año == 0 || periodo == 0) {
                Alert alerta = new Alert(Alert.AlertType.WARNING);
                alerta.setTitle("Periodo");
                alerta.setHeaderText(null);
                alerta.setContentText("Selecciona un año y periodo");
                alerta.showAndWait();
                return;
            }

            try {
                // Llamar al método readAllExcelFiles pasando todos los filtros
                data = readAllExcelFiles(año, periodo, tipoCapacitacion, departamento, acreditacion, nivel);
            } catch (InvalidFormatException ex) {
                Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
            }
            tabla.setItems(data);
            docentesTomandoCursos.setText("" + contarFilasUnicas(data));

        });

        botonExportar.setOnMouseClicked(event -> {
            String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\reporte.xlsx";

            switch (comboFormato.getValue() == null ? "default" : comboFormato.getValue()) {
                case "PDF":
                    exportarArchivo(rutaArchivo, (Stage) botonExportar.getScene().getWindow(), 1);
                    break;
                case "EXCEL":
                    exportarArchivo(rutaArchivo, (Stage) botonExportar.getScene().getWindow(), 2);
                    break;
                default:
            }
        });

        botonActualizarDatos.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.regresar(event, "ModificacionDatos", getClass());
            } catch (IOException ex) {
                Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

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

    public void actualizarDocumentos(MouseEvent event) throws IOException {
        ControladorGeneral.regresar(event, "ImportacionArchivos", getClass());
    }
}
