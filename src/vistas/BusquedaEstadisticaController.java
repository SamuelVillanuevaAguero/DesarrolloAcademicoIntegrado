package vistas;

import com.aspose.cells.SaveFormat;
import java.io.File;
import java.io.FileNotFoundException;
import java.net.URL;
import java.nio.file.Files;
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
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.MouseEvent;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import utilerias.busqueda.filaDato;
import utilerias.general.ControladorGeneral;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeParseException;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utilerias.busqueda.Docente;
import utilerias.busqueda.ExcelReader;

public class BusquedaEstadisticaController implements Initializable {

    //BOTONES
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

    //LISTAS DE OPCIONES (COMBOBOXES)
    @FXML
    private ComboBox<String> comboTipoCapacitacion;
    @FXML
    private ComboBox<String> comboDepartamento;
    @FXML
    private ComboBox<String> comboAcreditacion;
    @FXML
    private ComboBox<Integer> comboAño;
    @FXML
    private ComboBox<String> comboPeriodo;
    @FXML
    private ComboBox<String> comboFormato;

    //TABLA
    @FXML
    private TableView<filaDato> tabla;

    //COLUMNAS DE LA TABLA
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

    //ETIQUETAS DE ESTADISTICAS
    @FXML
    private Label numeroTotalDocentes;
    @FXML
    private Label docentesTomandoCursos;
    @FXML
    private Label porcentajeDocentesCapacitados;
    @FXML
    private Label NumeroDocentesNivel;

    //AUXILIAR PARA EL LLENADO DE LA TABLA
    private ObservableList data;

    //TOTALES DE DOCENTES ADSCRITOS A CADA DEPARTAMENTO
    Map<String, Integer> totalDepartamentos;
    private int totalDocentes;

    //MÉTODOS RELACIONADOS CON LOS TOTALES POR CADA DEPARTAMENTO ------------------------------------------------
    private void asignarTotales(String ruta, int año, int periodo) {
        totalDepartamentos = new HashMap<>();

        int numeroDepartamento = periodo == 1 ? 3 : 12;
        totalDocentes = obtenerTotales(ruta, año, numeroDepartamento);

        for (String depa : comboDepartamento.getItems()) {
            numeroDepartamento++;
            //System.out.println(numeroDepartamento + ":" + depa);
            totalDepartamentos.put(depa, obtenerTotales(ruta, año, numeroDepartamento));
        }
    }

    private int obtenerTotales(String ruta, int año, int departamento) {
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
            return (int) hoja.getRow(departamento).getCell(1).getNumericCellValue();

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
    //FIN--------------------------------------------------------------------------------------------------------

    //MÉTODOS DE BUSQUÉDA-----------------------------------------------------------------------------------------
    public void metodoBuscar() {
        String tipoCapacitacion = comboTipoCapacitacion.getValue() == null ? null : comboTipoCapacitacion.getValue().equals("Formación docente") ? "FD" : "AP";

        String departamento = comboDepartamento.getValue();
        String acreditacion = comboAcreditacion.getValue();
        String nivel = null;

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
        int filasDistintas = obtenerNumeroFilasUnicas(data);

        if (departamento != null) {
            int totalDocentes = totalDepartamentos.get(departamento);
            numeroTotalDocentes.setText("" + totalDocentes);
            porcentajeDocentesCapacitados.setText("" + (int) ((double) filasDistintas / totalDocentes * 100) + "%");
        }
        docentesTomandoCursos.setText("" + filasDistintas);
    }
    //FIN---------------------------------------------------------------------------------------------------------

    //MÉRTODOS DE EXPORTACIÓN-------------------------------------------------------------------------------------
    public void exportarAPDF(Workbook poiWorkbook, File archivoExportado) {
        try {
            // Guardar el Workbook de Apache POI en un archivo temporal
            File archivoTemporal = File.createTempFile("temp_excel", ".xlsx");
            try (FileOutputStream fos = new FileOutputStream(archivoTemporal)) {
                poiWorkbook.write(fos);
            }

            // Cargar el archivo temporal en un Workbook de Aspose.Cells
            com.aspose.cells.Workbook asposeWorkbook = new com.aspose.cells.Workbook(archivoTemporal.getAbsolutePath());

            // Exportar como PDF
            asposeWorkbook.save(archivoExportado.getAbsolutePath(), SaveFormat.PDF);

            System.out.println("Archivo exportado exitosamente como PDF en: " + archivoExportado.getAbsolutePath());

            // Eliminar el archivo temporal
            archivoTemporal.delete();
        } catch (Exception e) {
            System.err.println("Error al exportar el archivo a PDF:");
            e.printStackTrace();
        }
    }

    public void exportarAPDF(com.aspose.cells.Workbook workbook, File archivoExportado) {
        try {
            // Guardar el archivo Excel como PDF
            workbook.save(archivoExportado.getAbsolutePath(), SaveFormat.PDF);

            System.out.println("Archivo exportado exitosamente como PDF en: " + archivoExportado.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("Error al exportar el archivo a PDF:");
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

    public void exportarArchivo(String rutaArchivo, String rutaExportacion, int version, int formato) {
        // Validar la ruta del archivo original
        if (rutaArchivo == null || rutaArchivo.isEmpty()) {
            System.out.println("La ruta del archivo original es inválida.");
            return;
        }

        // Validar la ruta de exportación
        if (rutaExportacion == null || rutaExportacion.isEmpty()) {
            System.out.println("La ruta de exportación es inválida.");
            return;
        }

        File carpeta = new File(rutaExportacion);
        if (carpeta.exists()) {

        } else {
            if (carpeta.mkdir()) {
                System.out.println("Se creó la carpeta exitosamente");
            }
        }

        try (FileInputStream fileInputStream = new FileInputStream(rutaArchivo); Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Ajustar el nombre del archivo según el formato
            String extension = (formato == 1) ? ".pdf" : ".xlsx";

            File archivoExportado;
            if (comboDepartamento.getValue() == null) {
                archivoExportado = new File(rutaExportacion + "\\reporte_(Version_" + (version + 1) + ")" + extension);
            } else {
                archivoExportado = new File(rutaExportacion + "\\reporte_(Version_" + (version + 1) + ")" + extension);
            }

            // Modificar el contenido del archivo si el formato es Excel
            if (formato == 2) {
                int periodo = comboPeriodo.getValue().equalsIgnoreCase("Enero - Julio") ? 1 : 2;
                //llenarExcel(workbook, tabla, comboAño.getValue(), periodo);
                Map<String, Docente> docenteMap = ExcelReader.readExcel(comboAño.getValue(), periodo);
                ExcelReader.processExcel(docenteMap, comboAño.getValue(), periodo);
                ExcelReader.writeToExcel(workbook, docenteMap, comboDepartamento.getValue(), comboAño.getValue(), periodo);

                // Guardar los cambios en el archivo exportado
                try (FileOutputStream outputStream = new FileOutputStream(archivoExportado)) {
                    workbook.write(outputStream); // Guardar los cambios en el archivo de destino
                }

                String mensaje = comboDepartamento.getValue() == null
                        ? "Reporte de todos los departamentos exportado exitosamente"
                        : "Reporte del departamento " + comboDepartamento.getValue() + " exportado exitosamente.";

                Alert alerta = new Alert(Alert.AlertType.INFORMATION);
                alerta.setTitle("Guardado EXCEL");
                alerta.setHeaderText(null);
                alerta.setContentText(mensaje);
                alerta.showAndWait();

                System.out.println("Archivo exportado exitosamente como Excel en: " + archivoExportado.getAbsolutePath());
            } else if (formato == 1) {
                // Exportar como PDF (implementación pendiente)
                int periodo = comboPeriodo.getValue().equalsIgnoreCase("Enero - Julio") ? 1 : 2;

                Map<String, Docente> docenteMap = ExcelReader.readExcel(comboAño.getValue(), periodo);
                ExcelReader.processExcel(docenteMap, comboAño.getValue(), periodo);
                ExcelReader.writeToExcel(workbook, docenteMap, comboDepartamento.getValue(), comboAño.getValue(), periodo);
                exportarAPDF(workbook, archivoExportado);

                String mensaje = comboDepartamento.getValue() == null
                        ? "Reporte de todos los departamentos exportado exitosamente"
                        : "Reporte del departamento " + comboDepartamento.getValue() + " exportado exitosamente.";

                Alert alerta = new Alert(Alert.AlertType.INFORMATION);
                alerta.setTitle("Guardado PDF");
                alerta.setHeaderText(null);
                alerta.setContentText(mensaje);
                alerta.showAndWait();

                System.out.println("Archivo exportado exitosamente como PDF en: " + archivoExportado.getAbsolutePath());
            } else {
                System.out.println("Formato no soportado.");
            }

        } catch (IOException e) {
            System.out.println("Error durante la exportación: " + e.getMessage());
            e.printStackTrace();
        }
    }
    //FIN---------------------------------------------------------------------------------------------------------

    //MÉTODOS DE ESCRITURA EN EXCEL-------------------------------------------------------------------------------
    public void llenarExcel(Workbook workbook, TableView<filaDato> tablaJavaFX, int año, int periodo) {
        try {
            CellStyle estilo = crearEstiloCelda(workbook);
            Sheet sheet = workbook.getSheetAt(0); // Obtener la primera hoja
            int filaInicio = 10; // Fila donde inicia la tabla en el Excel (basado en tu imagen)
            int totalAP = 0;
            int totalFD = 0;

            // Obtener los datos de la tabla de JavaFX
            ObservableList<filaDato> listaDocentes = tablaJavaFX.getItems();

            //Datos generales
            Row fila = sheet.getRow(4);
            if (fila == null) {
                fila = sheet.createRow(4);
            }
            Cell celda = fila.createCell(3);
            celda.setCellValue(comboDepartamento.getValue().toUpperCase());
            celda.setCellStyle(estiloNegrillas(workbook));

            //Año
            fila = sheet.getRow(5);
            if (fila == null) {
                fila = sheet.createRow(5);
            }
            celda = fila.createCell(3);
            celda.setCellValue("AÑO " + comboAño.getValue().toString().toUpperCase());
            celda.setCellStyle(estiloNegrillas(workbook));

            fila = sheet.getRow(7);
            if (fila == null) {
                fila = sheet.createRow(7);
            }
            celda = fila.createCell(3);
            celda.setCellValue(periodo == 1 ? "ENERO - JULIO" : "AGOSTO - DICIEMBRE");
            celda.setCellStyle(estiloPeriodo(workbook));

            // Llenar las filas del archivo Excel
            int y = 0;

            Map<String, Integer> mapaDocente = new HashMap<>();

            for (int i = 0; i < listaDocentes.size(); i++) {

                filaDato filaDato = listaDocentes.get(i);

                // Obtener o crear la fila en el Excel
                fila = sheet.getRow(filaInicio + y);
                if (fila == null) {
                    fila = sheet.createRow(filaInicio + y);
                }

                if (filaDato.getAcreditado().equalsIgnoreCase("si")) {
                    // Escribir las celdas

                    if (mapaDocente.containsKey(filaDato.getNombre() + " " + filaDato.getApellidoPaterno() + " " + filaDato.getApellidoMaterno())) {
                        System.out.println("YA EXISTE");
                        int row = mapaDocente.get(filaDato.getNombre() + " " + filaDato.getApellidoPaterno() + " " + filaDato.getApellidoMaterno());
                        fila = sheet.getRow(row);
                        // Verificar el tipo de capacitación para asignar a la columna adecuada
                        if (filaDato.getTipoCapacitacion().equalsIgnoreCase("FD")) {
                            totalFD += filaDato.getNoCursos();
                            celda = fila.createCell(3);
                            celda.setCellStyle(estilo);
                            celda.setCellValue(filaDato.getNoCursos()); // Columna C: Número de curso (AP)
                        } else {
                            totalAP += filaDato.getNoCursos();
                            celda = fila.createCell(4);
                            celda.setCellStyle(estilo);
                            celda.setCellValue(filaDato.getNoCursos()); // Columna D: Número de curso (FD)
                        }
                    } else {
                        System.out.println("NO EXISTE");
                        mapaDocente.put(filaDato.getNombre() + " " + filaDato.getApellidoPaterno() + " " + filaDato.getApellidoMaterno(), filaInicio + y);
                        y++;
                        celda = fila.createCell(1);
                        celda.setCellStyle(estilo);
                        celda.setCellValue(y); // Columna A: Número (índice)

                        celda = fila.createCell(2);
                        celda.setCellStyle(estilo);
                        celda.setCellValue(filaDato.getNombre() + " " + filaDato.getApellidoPaterno() + " " + filaDato.getApellidoMaterno()); // Columna B: Nombre del docente

                        // Verificar el tipo de capacitación para asignar a la columna adecuada
                        if (filaDato.getTipoCapacitacion().equalsIgnoreCase("FD")) {
                            totalFD += filaDato.getNoCursos();
                            celda = fila.createCell(3);
                            celda.setCellStyle(estilo);
                            celda.setCellValue(filaDato.getNoCursos()); // Columna C: Número de curso (AP)

                            celda = fila.createCell(4);
                            celda.setCellStyle(estilo);
                            celda.setCellValue(""); // Columna C: Número de curso (AP)
                        } else {
                            totalAP += filaDato.getNoCursos();
                            celda = fila.createCell(3);
                            celda.setCellStyle(estilo);
                            celda.setCellValue(""); // Columna D: Número de curso (FD)

                            celda = fila.createCell(4);
                            celda.setCellStyle(estilo);
                            celda.setCellValue(filaDato.getNoCursos()); // Columna D: Número de curso (FD)
                        }
                    }

                }

            }

            fila = sheet.getRow(filaInicio + y);
            if (fila == null) {
                fila = sheet.createRow(filaInicio + y);
            }

            celda = fila.createCell(2);
            celda.setCellStyle(estilo);
            celda.setCellValue("TOTAL");

            celda = fila.createCell(3);
            celda.setCellStyle(estilo);
            celda.setCellValue(totalFD);

            celda = fila.createCell(4);
            celda.setCellStyle(estilo);
            celda.setCellValue(totalAP);

            fila = sheet.getRow(filaInicio + y + 3);
            if (fila == null) {
                fila = sheet.createRow(filaInicio + y + 3);
            }

            celda = fila.createCell(3);
            celda.setCellStyle(estiloNegrillas(workbook));
            celda.setCellValue(obtenerJefe(ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx", año, periodo).toUpperCase());

            // Combinar las celdas desde (fila 0, columna 0) hasta (fila 0, columna 3)
            CellRangeAddress rango = new CellRangeAddress(filaInicio + y + 3, filaInicio + y + 3, 3, 4); // Fila inicial, Fila final, Columna inicial, Columna final
            sheet.addMergedRegion(rango);

            fila = sheet.getRow(filaInicio + y + 4);
            if (fila == null) {
                fila = sheet.createRow(filaInicio + y + 4);
            }

            celda = fila.createCell(3);
            celda.setCellStyle(estiloNegrillas(workbook));
            celda.setCellValue("JEFE DEL DEPARTAMENTO DE DESARROLLO ACADEMICO");

            // Combinar las celdas desde (fila 0, columna 0) hasta (fila 0, columna 3)
            rango = new CellRangeAddress(filaInicio + y + 4, filaInicio + y + 4, 3, 4); // Fila inicial, Fila final, Columna inicial, Columna final
            sheet.addMergedRegion(rango);

            System.out.println("Archivo Excel actualizado correctamente.");

        } catch (Exception e) {
            System.err.println("Error al llenar el archivo Excel: " + e.getMessage());
            e.printStackTrace();
        }
    }
    //FIN---------------------------------------------------------------------------------------------------------

    //MÉTODOS DE VERSIONES---------------------------------------------------------------------------------------
    public int obtenerUltimaSemana(String carpetaDestino, String nombreArchivo, String versionS, String extension) {
        File carpeta = new File(carpetaDestino);

        // Validar que la carpeta existe y es un directorio
        if (!carpeta.exists() || !carpeta.isDirectory()) {
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

        return maxVersion;
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
    //FIN--------------------------------------------------------------------------------------------------------

    //MÉTODOS DE ESTILOS EXCEL----------------------------------------------------------------------------------
    private CellStyle crearEstiloCelda(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        return style;
    }

    private CellStyle estiloNegrillas(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // Crear fuente en negritas
        Font font = workbook.createFont();
        font.setBold(true); // Hacer la fuente en negrita
        style.setFont(font);

        // Configurar alineación centrada
        style.setAlignment(CellStyle.ALIGN_CENTER); // Centrar horizontalmente
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); // Centrar verticalmente

        // Configurar bordes
        return style;
    }

    private CellStyle estiloPeriodo(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // Crear fuente en negritas
        Font font = workbook.createFont();
        font.setBold(true); // Hacer la fuente en negrita
        style.setFont(font);

        // Configurar alineación centrada
        style.setAlignment(CellStyle.ALIGN_CENTER); // Centrar horizontalmente
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); // Centrar verticalmente

        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);

        // Configurar bordes
        return style;
    }
    //FIN-------------------------------------------------------------------------------------------------------

    //MÉTODOS DE OBTENER----------------------------------------------------------------------------------------
    public int obtenerMes(String fecha) {
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

    public static int obtenerAnio(String fecha) {
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

    public int obtenerNumeroFilasUnicas(ObservableList<filaDato> data) {
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

    private String obtenerCadenaCelda(Cell cell) {
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

    public List<Integer> obtenerAniosDisponibles(String baseDirectoryPath) {
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

    private String obtenerJefe(String ruta, int año, int periodo) {
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
            return hoja.getRow(1).getCell(1).getStringCellValue();

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
    //FIN-------------------------------------------------------------------------------------------------------

    //MÉTODO PARA LLENAR LA TABLA-------------------------------------------------------------------------------
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
                            currentPeriodo = obtenerMes(periodoCell) < 7 ? 1 : 2; // Determinar el período
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
                    String filaTipoCapacitacion = obtenerCadenaCelda(row.getCell(8));
                    String filaDepartamento = obtenerCadenaCelda(row.getCell(5));
                    String filaPosgrado = obtenerCadenaCelda(row.getCell(5)).equalsIgnoreCase("POSGRADO") ? "si" : "no";
                    String filaAcreditacion = obtenerCadenaCelda(row.getCell(11));

                    // Validar filtros de ComboBox
                    boolean passTipoCapacitacion = tipoCapacitacion == null || tipoCapacitacion.equalsIgnoreCase(filaTipoCapacitacion);
                    boolean passDepartamento = departamento == null || departamento.equalsIgnoreCase(filaDepartamento);
                    boolean passAcreditacion = acreditacion == null || acreditacion.equalsIgnoreCase(filaAcreditacion) || acreditacion.equalsIgnoreCase("Ambos");
                    boolean passNivel = nivel == null
                            || (nivel.equals("Licenciatura") && filaPosgrado.equalsIgnoreCase("No"))
                            || (nivel.equals("Posgrado") && filaPosgrado.equalsIgnoreCase("Si"));

                    // Aplicar todos los filtros
                    if (passAñoFilter && passPeriodoFilter && passTipoCapacitacion && passDepartamento && passAcreditacion && passNivel) {
                        String nombre = obtenerCadenaCelda(row.getCell(2));
                        String apellidoPaterno = obtenerCadenaCelda(row.getCell(0));
                        String apellidoMaterno = obtenerCadenaCelda(row.getCell(1));

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
    //FIN-------------------------------------------------------------------------------------------------------

//----------------------------------------------------------MÉTODOS DE LA INTERFAZ GRAFICA------------------------------------------------------------------------------------------//
    //MÉTODO INICIAL--------------------------------------------------------------------------------------------
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

        // Inicializar opciones de ComboBoxes
        comboTipoCapacitacion.getItems().addAll("Actualización profesional", "Formación docente");
        comboDepartamento.getItems().addAll("CIENCIAS BÁSICAS", "CIENCIAS ECONÓMICO ADMINISTRATIVAS", "CIENCIAS DE LA TIERRA", "INGENIERÍA INDUSTRIAL", "METAL MECÁNICA", "QUÍMICA Y BIOQUÍMICA", "SISTEMAS COMPUTACIONALES", "POSGRADO");
        comboAcreditacion.getItems().addAll("Si", "No", "Ambos");

        tabla.setItems(data);
        int docentesCursos = obtenerNumeroFilasUnicas(data);

        Calendar calendario = Calendar.getInstance();
        int year = calendario.get(Calendar.YEAR);
        int mesActual = calendario.get(Calendar.MONTH) + 1;
        asignarTotales(ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx", year, mesActual < 7 ? 1 : 2);

        double porcentajeCapacitados = (double) docentesCursos / totalDocentes * 100;

        docentesTomandoCursos.setText(String.valueOf(docentesCursos));
        numeroTotalDocentes.setText(String.valueOf(totalDocentes));

        porcentajeDocentesCapacitados.setText((int) porcentajeCapacitados + "%");

        int currentYear = Year.now().getValue();

        comboAño.getItems().addAll(obtenerAniosDisponibles(ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos"));
        comboPeriodo.getItems().addAll("Enero - Julio", "Agosto - Diciembre");
        comboFormato.getItems().addAll("PDF", "EXCEL");

        comboAño.setValue(currentYear);
        comboPeriodo.setValue(mesActual < 7 ? "Enero - Julio" : "Agosto - Diciembre");

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
            comboFormato.setValue(null);
            tabla.getItems().clear();
            metodoBuscar();

            int filasDistintas = obtenerNumeroFilasUnicas(data);

            numeroTotalDocentes.setText("" + totalDocentes);
            porcentajeDocentesCapacitados.setText("" + (int) ((double) filasDistintas / totalDocentes * 100) + "%");

            docentesTomandoCursos.setText("" + filasDistintas);
        });

        botonBuscar.setOnMouseClicked(event -> {

            System.out.println("Botón buscar...");
            metodoBuscar();

        });

        botonExportar.setOnMouseClicked(event -> {

            if (comboFormato.getValue() == null) {
                Alert alerta = new Alert(Alert.AlertType.WARNING);
                alerta.setTitle("Formato");
                alerta.setHeaderText(null);
                alerta.setContentText("Selecciona un formato de exportación");
                alerta.showAndWait();
                return;
            }
            /*if (comboDepartamento.getValue() == null) {
                Alert alerta = new Alert(Alert.AlertType.WARNING);
                alerta.setTitle("Departamento");
                alerta.setHeaderText(null);
                alerta.setContentText("Selecciona departamento para exportar");
                alerta.showAndWait();
                return;
            }*/

            metodoBuscar();

            int año = comboAño.getValue();
            int periodo = comboPeriodo.getValue().equalsIgnoreCase("Enero - Julio") ? 1 : 2;
            String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_importados\\" + año + "\\" + periodo + "-" + año + "\\formato_de_reporte_para_docentes_capacitados\\";
            int version = obtenerUltimaSemana(rutaArchivo, "formato\\_\\(Version_\\d+\\)\\.xlsx", "Version", "xlsx");
            rutaArchivo += "formato_(Version_" + version + ").xlsx";

            String rutaExportacion = "";
            if (comboDepartamento.getValue() == null) {
                rutaExportacion = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_exportados\\" + año + "\\" + periodo + "-" + año + "\\reportes_estadisticos\\TODOS";
            } else {
                rutaExportacion = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_exportados\\" + año + "\\" + periodo + "-" + año + "\\reportes_estadisticos" + "\\" + comboDepartamento.getValue();
            }

            System.out.println("REXPORTACIÓN:" + rutaExportacion);
            int versionReporte = obtenerUltimaSemana(rutaExportacion, "reporte\\_\\(Version_\\d+\\)\\.xlsx", "Version", "xlsx");
            switch (comboFormato.getValue() == null ? "default" : comboFormato.getValue()) {
                case "PDF":
                    versionReporte = obtenerUltimaSemana(rutaExportacion, "reporte\\_\\(Version_\\d+\\)\\.pdf", "Version", "pdf");
                    exportarArchivo(rutaArchivo, rutaExportacion, versionReporte, 1);
                    break;
                case "EXCEL":
                    exportarArchivo(rutaArchivo, rutaExportacion, versionReporte, 2);
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
    //FIN-------------------------------------------------------------------------------------------------------

    //OPCIONES DE LA BARRA SUPERIOR ----------------------------------------------------------------------------
    public void cerrarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión?", getClass());
    }

    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    public void regresarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.regresar(event, "Principal", getClass());
    }
    //FIN---------------------------------------------------------------------------------------------------------

    //OPCIÓN DE REDIRECCIÓN A LA VENTANA DE IMPORTACIÓN DE DATOS--------------------------------------------------
    public void actualizarDocumentos(MouseEvent event) throws IOException {
        ControladorGeneral.regresar(event, "ImportacionArchivos", getClass());
    }
    //FIN---------------------------------------------------------------------------------------------------------

    //MÉTODO PARA LEER EL EXCEL DE LOS DOCENTES ADSCTRITOS
}
