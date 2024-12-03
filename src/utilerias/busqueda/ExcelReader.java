/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package utilerias.busqueda;

/**
 *
 * @author Samue
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utilerias.general.ControladorGeneral;
import vistas.BusquedaEstadisticaController;

public class ExcelReader {

    public static Map<String, Docente> readExcel(int año, int periodo) {
        String filePath = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_importados\\" + año + "\\" + periodo + "-" + año + "\\listado_de_docentes_adscritos\\";
        int version = obtenerUltimaSemana(filePath, "listado\\_\\(Semana_\\d+\\)\\.xlsx", "Semana", "xlsx");
        filePath += "listado_(Semana_" + version + ").xlsx";
        Map<String, Docente> docenteMap = new HashMap<>();
        try (Workbook workbook = new XSSFWorkbook(filePath)) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String departamento = sheet.getSheetName();
                for (Row row : sheet) {
                    if (row.getRowNum() > 0) { //Empezar después de los encabezados
                        //3 celdas para el nombre completo jeje
                        String nombre = getCellValueAsString(row.getCell(2)).trim()
                                + " " + getCellValueAsString(row.getCell(1)).trim()
                                + " " + getCellValueAsString(row.getCell(0)).trim();

                        if (nombre != null && !nombre.isEmpty()) {
                            Docente docente = new Docente(
                                    0, // Contador de docentes, no supe cómo ponerle mmmm, lo quito?
                                    nombre,
                                    0, // Contador de cursos FD
                                    0, // COntador de cursos AP, no es apellido paterno JAJAJAJA
                                    departamento
                            );
                            docenteMap.put(nombre, docente);
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return docenteMap;
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            default:
                return null;
        }
    }

    public static void processExcel(Map<String, Docente> docenteMap, int año, int periodo) {

        String filePath = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos\\" + año + "\\" + periodo + "-" + año + "\\";
        int version = obtenerUltimaSemana(filePath, "condensado\\_\\(version_\\d+\\)\\.xlsx", "version", "xlsx");
        filePath += "condensado_(version_" + version + ").xlsx";

        try (Workbook workbook = new XSSFWorkbook(filePath)) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                for (Row row : sheet) {
                    if (row.getRowNum() > 0) { // Ignore header row
                        String nombre = getCellValueAsString(row.getCell(2)).trim() + " "
                                + getCellValueAsString(row.getCell(1)).trim() + " "
                                + getCellValueAsString(row.getCell(0)).trim();

                        if (nombre != null && !nombre.isEmpty()) {
                            Docente docente = docenteMap.get(nombre);
                            if (docente != null) {
                                String acreditado = getCellValueAsString(row.getCell(11));
                                if (acreditado != null && acreditado.equalsIgnoreCase("si")) {
                                    String nombreEvento = getCellValueAsString(row.getCell(7));
                                    int numeroCurso = Integer.parseInt(nombreEvento.replaceAll("[^0-9]", ""));
                                    docente.getListaCursos().put(numeroCurso, nombreEvento);

                                    String capacitacion = getCellValueAsString(row.getCell(8));
                                    if (capacitacion != null && capacitacion.equalsIgnoreCase("FD")) {
                                        docente.setNumeroCursosFD(docente.getNumeroCursosFD() + 1);
                                    } else if (capacitacion != null && capacitacion.equalsIgnoreCase("AP")) {
                                        docente.setNumeroCursosAP(docente.getNumeroCursosAP() + 1);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void writeToExcel(Workbook workbook, Map<String, Docente> docenteMap, String departamento, int año, int periodo) {
        try {
            CellStyle estiloCeldas = crearEstiloCelda(workbook);
            Sheet sheet = workbook.getSheetAt(0);

            int no = 1;
            int rowIndex = 10;
            int totalAP = 0;
            int totalFD = 0;

            Row rowDatosEncabezado = sheet.createRow(4);
            Cell celdaEncabezado = rowDatosEncabezado.createCell(4);
            celdaEncabezado.setCellValue(departamento == null ? "TODOS LOS DEPARTAMENTOS" : departamento);
            celdaEncabezado.setCellStyle(estiloNegrillas(workbook));

            rowDatosEncabezado = sheet.createRow(5);
            celdaEncabezado = rowDatosEncabezado.createCell(4);
            celdaEncabezado.setCellValue("AÑO " + año);
            celdaEncabezado.setCellStyle(estiloNegrillas(workbook));

            rowDatosEncabezado = sheet.createRow(7);
            celdaEncabezado = rowDatosEncabezado.createCell(4);
            celdaEncabezado.setCellValue(periodo == 1 ? "ENERO - JULIO " + año : "AGOSTO - DICIEMBRE " + año);
            celdaEncabezado.setCellStyle(estiloPeriodo(workbook));

            celdaEncabezado = rowDatosEncabezado.createCell(5);
            celdaEncabezado.setCellStyle(estiloPeriodo(workbook));

            for (Map.Entry<String, Docente> entry : docenteMap.entrySet()) {
                Docente docente = entry.getValue();
                if (departamento == null || docente.getDepartamento().equalsIgnoreCase(departamento)) {
                    Row row = sheet.createRow(rowIndex);

                    int numeroCursos = docente.getListaCursos().size();

                    // NO.              
                    if (numeroCursos > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 1, 1));
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 1, 1));
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 1, 1));
                    }

                    Cell celda = row.createCell(1);
                    celda.setCellValue(no);
                    no++;
                    celda.setCellStyle(estiloCeldas);

                    if (numeroCursos > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 2, 2));
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 2, 2));
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 2, 2));
                    }

                    // NOMBRE DEL DOCENTE
                    celda = row.createCell(2);
                    celda.setCellValue(docente.getNombre());
                    celda.setCellStyle(estiloCeldas);

                    // Códigos
                    //row.createCell(2).setCellValue(codigosBuilder.toString().trim());
                    // F. DOCENTE
                    if (numeroCursos > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 4, 4));
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 4, 4));
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 4, 4));
                    }

                    celda = row.createCell(4);
                    celda.setCellValue(docente.getNumeroCursosFD());
                    celda.setCellStyle(estiloCeldas);

                    totalFD += docente.getNumeroCursosFD();
                    totalAP += docente.getNumeroCursosAP();

                    // A. PROFESIONAL
                    if (numeroCursos > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 5, 5));
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 5, 5));
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + numeroCursos - 1, 5, 5));
                    }

                    celda = row.createCell(5);
                    celda.setCellValue(docente.getNumeroCursosAP());
                    celda.setCellStyle(estiloCeldas);

                    row.createCell(3);

                    int nRow = rowIndex;
                    StringBuilder codigosBuilder = new StringBuilder();
                    for (Map.Entry<Integer, String> curso : docente.getListaCursos().entrySet()) {
                        Row rowAuxiliar = sheet.getRow(nRow);
                        if (rowAuxiliar == null) {
                            rowAuxiliar = sheet.createRow(nRow);
                        }
                        celda = rowAuxiliar.createCell(3);
                        celda.setCellValue(curso.getKey());
                        celda.setCellStyle(estiloCeldas);

                        celda = rowAuxiliar.getCell(1);
                        if (celda == null) {
                            celda = rowAuxiliar.createCell(1);
                        }
                        celda.setCellStyle(estiloCeldas);

                        celda = rowAuxiliar.getCell(5);
                        if (celda == null) {
                            celda = rowAuxiliar.createCell(5);
                        }
                        celda.setCellStyle(estiloCeldas);
                        nRow++;
                    }

                    rowIndex = numeroCursos == 0 ? rowIndex + 1 : rowIndex + numeroCursos;
                }

            }

            Row row = sheet.createRow(rowIndex);
            Cell celda = row.createCell(4);
            //celda.setCellFormula("SUM(E" + 10 + ":E" + (rowIndex) + ")");
            celda.setCellValue(totalFD);
            celda.setCellStyle(estiloPeriodo(workbook));

            celda = row.createCell(5);
            //celda.setCellFormula("SUM(F" + 10 + ":F" + (rowIndex) + ")");
            celda.setCellValue(totalAP);
            celda.setCellStyle(estiloPeriodo(workbook));

            celda = row.createCell(3);
            celda.setCellValue("TOTAL");
            celda.setCellStyle(estiloPeriodo(workbook));

            //INFORMACIÓN DEL JEFE DE DESARROLLO ACADEMICO
            row = sheet.createRow(rowIndex + 4);
            celda = row.createCell(4);
            celda.setCellValue(obtenerJefe(ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx", año, periodo).toUpperCase());
            celda.setCellStyle(estiloNegrillas(workbook));

            celda = row.createCell(5);
            celda.setCellStyle(estiloNegrillas(workbook));

            CellRangeAddress rango = new CellRangeAddress(rowIndex + 4, rowIndex + 4, 4, 5);
            sheet.addMergedRegion(rango);

            //TEXTO DEL JEFE
            row = sheet.createRow(rowIndex + 5);
            celda = row.createCell(4);
            celda.setCellValue("JEFE DEL DEPARTAMENTO DE DESARROLLO ACADEMICO");
            celda.setCellStyle(estiloNegrillas(workbook));

            celda = row.createCell(5);
            celda.setCellStyle(estiloNegrillas(workbook));

            rango = new CellRangeAddress(rowIndex + 5, rowIndex + 5, 4, 5);
            sheet.addMergedRegion(rango);

            // Ajustar ancho de columnas
            for (int i = 0; i <= 5; i++) {
                if (i == 4 || i == 5) {
                    sheet.setColumnWidth(i, 8000);
                } else {
                    sheet.autoSizeColumn(i);
                }
            }

            // Guardar archivo
            /*try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                workbook.write(outputStream);
            }*/
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //ESTILOS
    private static CellStyle crearEstiloCelda(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER); // Centrar horizontalmente
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); // Centrar verticalmente
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        return style;
    }

    private static CellStyle estiloNegrillas(Workbook workbook) {
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

    private static CellStyle estiloPeriodo(Workbook workbook) {
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

    //OBTENER VSERIÓN
    private static int obtenerUltimaSemana(String carpetaDestino, String nombreArchivo, String versionS, String extension) {
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
                    System.out.println("Algo salió mal al encontrar la ultima versión del archivo de docentes adscritos");
                }
            }
        }

        return maxVersion;
    }

    //OBTENER JEFE DEL DEPARTAMENTO DE DESARROLLO ACADEMICO
    private static String obtenerJefe(String ruta, int año, int periodo) {
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
}
