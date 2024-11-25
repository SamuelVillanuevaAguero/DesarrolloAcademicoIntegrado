package utilerias.principal;

import static com.aspose.cells.PropertyType.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.io.File;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import utilerias.general.ControladorGeneral;

public class Necesidadcapacitacion {

    private String obtenerUltimaSemana(String rutaDirectorio, String patron, String identificador) {
        File directorio = new File(rutaDirectorio);
        File[] archivos = directorio.listFiles();
        int ultimaSemana = 0;
        
        if (archivos != null) {
            Pattern pattern = Pattern.compile(patron);
            for (File archivo : archivos) {
                Matcher matcher = pattern.matcher(archivo.getName());
                if (matcher.find()) {
                    String numeroStr = archivo.getName().split(identificador + "_")[1].split("\\)")[0];
                    int numero = Integer.parseInt(numeroStr);
                    if (numero > ultimaSemana) {
                        ultimaSemana = numero;
                    }
                }
            }
        }
        return String.valueOf(ultimaSemana);
    }

    
    public void generarArchivo() throws IOException {
        // Obtener año y periodo actual
        Calendar calendario = Calendar.getInstance();
        int año = calendario.get(Calendar.YEAR);
        int periodo = (calendario.get(Calendar.MONTH) + 1) < 7 ? 1 : 2;

        // Construir rutas base para archivos de entrada
        String rutaBaseEntrada = ControladorGeneral.obtenerRutaDeEjecusion() + 
                         "\\Gestion_de_Cursos\\Archivos_importados\\" + 
                         año + "\\" + periodo + "-" + año + "\\";

        // Construir ruta base para archivo de salida
        String rutaBaseSalida = ControladorGeneral.obtenerRutaDeEjecusion() + 
                         "\\Gestion_de_Cursos\\Sistema\\informacion_notificaciones\\" +
                         año + "\\" + periodo + "-" + año + "\\";

        // Construir rutas específicas de entrada
        String rutaCapacitacionDir = rutaBaseEntrada + "listado_de_pre_regitro_a_cursos_de_capacitacion\\";
        String rutaNecesidadesDir = rutaBaseEntrada + "listado_de_deteccion_de_necesidades\\";

        // Obtener última semana para cada archivo
        String semanaCapacitacion = obtenerUltimaSemana(rutaCapacitacionDir, 
            "listado\\_\\(Semana_\\d+\\)\\.xlsx", "Semana");
        String semanaNecesidades = obtenerUltimaSemana(rutaNecesidadesDir, 
            "listado\\_\\(Semana_\\d+\\)\\.xlsx", "Semana");

        // Construir rutas completas de los archivos de entrada
        String rutaCapacitacion = rutaCapacitacionDir + 
            "listado_(Semana_" + semanaCapacitacion + ").xlsx";
        String rutaNecesidades = rutaNecesidadesDir + 
            "listado_(Semana_" + semanaNecesidades + ").xlsx";

        // Crear estructura de directorios para la salida
        File directorioSalida = new File(rutaBaseSalida);
        if (!directorioSalida.exists()) {
            if (!directorioSalida.mkdirs()) {
                throw new IOException("No se pudieron crear los directorios necesarios para la salida");
            }
        }

        // Construir ruta completa del archivo de salida
        String rutaSalida = rutaBaseSalida + "docentes_recomendables_(Semana_" + 
            semanaCapacitacion + ").xlsx";

        // Leer archivos
        List<String[]> listaCapacitacion = leerCapacitacion(rutaCapacitacion);
        Map<String, String[]> necesidades = leerNecesidades(rutaNecesidades);

        // Generar nuevo archivo
        generarArchivoNuevo(listaCapacitacion, necesidades, rutaSalida);

        System.out.println("Archivo generado correctamente en: " + rutaSalida);
    }

    private static List<String[]> leerCapacitacion(String rutaArchivo) throws IOException {
        List<String[]> profesores = new ArrayList<>();
        FileInputStream fis = new FileInputStream(rutaArchivo);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue; // Saltar encabezado
            }
            String apellidoPaterno = leerCeldaComoTexto(row.getCell(4)); // Columna E
            String apellidoMaterno = leerCeldaComoTexto(row.getCell(5)); // Columna F
            String nombre = leerCeldaComoTexto(row.getCell(6)); // Columna G

            if (!nombre.isEmpty()) {
                profesores.add(new String[]{nombre, apellidoPaterno, apellidoMaterno});
            }
        }

        // Ordenar por nombre, apellido paterno, apellido materno
        profesores.sort(Comparator.comparing((String[] p) -> p[0])
                .thenComparing(p -> p[1])
                .thenComparing(p -> p[2]));

        workbook.close();
        fis.close();
        return profesores;
    }

    private static Map<String, String[]> leerNecesidades(String rutaArchivo) throws IOException {
        Map<String, String[]> necesidades = new HashMap<>();
        FileInputStream fis = new FileInputStream(rutaArchivo);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // Recorrer filas del archivo Excel
        for (Row row : sheet) {
            if (row.getRowNum() < 2) { // Saltar encabezados (las primeras dos filas)
                continue;
            }

            // Leer valores de las celdas
            String nombreCompleto = leerCeldaComoTexto(row.getCell(0)); // Columna A: Nombre Completo
            String fd = leerCeldaComoTexto(row.getCell(2)); // Columna C: FD 
            String ap = leerCeldaComoTexto(row.getCell(3)); // Columna D: AP 

            // Verificar si FD o AP contiene "Recomendable"
            if (!nombreCompleto.isEmpty()
                    && ("Recomendable".equalsIgnoreCase(fd) || "Recomendable".equalsIgnoreCase(ap))) {
                necesidades.put(nombreCompleto, new String[]{
                    "Recomendable".equalsIgnoreCase(fd) ? "Recomendable" : "",
                    "Recomendable".equalsIgnoreCase(ap) ? "Recomendable" : ""
                });
            }
        }

        workbook.close();
        fis.close();
        return necesidades;
    }

    private static void generarArchivoNuevo(List<String[]> listaCapacitacion, Map<String, String[]> necesidades, String rutaSalida) throws IOException {
        Workbook workbookNuevo = new XSSFWorkbook();
        Sheet sheetNuevo = workbookNuevo.createSheet("Docentes Recomendables");

        // Crear encabezados
        Row headerRow1 = sheetNuevo.createRow(0); // Primera fila para el título
        headerRow1.createCell(0).setCellValue("Nombre");
        headerRow1.createCell(1).setCellValue("Necesidad de capacitación detectada");
        headerRow1.createCell(2).setCellValue("Necesidad de capacitación detectada");

        Row headerRow2 = sheetNuevo.createRow(1); // Segunda fila para subcategorías
        headerRow2.createCell(1).setCellValue("FD"); 
        headerRow2.createCell(2).setCellValue("AP"); 

        // Crear un conjunto de nombres completos en el archivo programa_de_capacitacion(1)
        Set<String> nombresCapacitacion = new HashSet<>();
        for (String[] profesor : listaCapacitacion) {
            String nombreCompleto = profesor[0] + " " + profesor[1] + " " + profesor[2];
            nombresCapacitacion.add(nombreCompleto.trim());
        }

        // Llenar datos en el nuevo archivo
        int rowIndex = 2; // Iniciar en la tercera fila
        for (Map.Entry<String, String[]> entry : necesidades.entrySet()) {
            String nombre = entry.getKey().trim(); // Nombre completo del docente
            String[] necesidad = entry.getValue(); // Necesidades FD y AP

            // Verificar si el nombre no está en programa_de_capacitacion(1)
            if (!nombresCapacitacion.contains(nombre)) {
                Row row = sheetNuevo.createRow(rowIndex++);

                // Escribir datos en las celdas
                row.createCell(0).setCellValue(nombre);     // Nombre
                row.createCell(1).setCellValue(necesidad[0]); // FD 
                row.createCell(2).setCellValue(necesidad[1]); // AP 
            }
        }

        // Guardar el archivo
        try (FileOutputStream fos = new FileOutputStream(rutaSalida)) {
            workbookNuevo.write(fos);
        }
        workbookNuevo.close();
    }

    // Método auxiliar para leer una celda como texto
    private static String leerCeldaComoTexto(Cell cell) {
        if (cell == null) {
            return ""; // Celda vacía
        }
        switch (cell.getCellType()) {
            case 1:
                return cell.getStringCellValue().trim();
            case 2:
                return String.valueOf(cell.getNumericCellValue()).trim();
            case 3:
                return String.valueOf(cell.getBooleanCellValue()).trim();
            case 4:
                return cell.getCellFormula().trim();
            default:
                return ""; // Otros tipos (BLANK, ERROR, etc.)
        }
    }
}
