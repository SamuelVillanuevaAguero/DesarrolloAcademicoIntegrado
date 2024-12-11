/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/javafx/FXMLController.java to edit this template
 */
package vistas;

import utilerias.principal.Necesidadcapacitacion;
import javafx.scene.layout.VBox;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.ResourceBundle;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Control;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
//import javafx.scene.control.*;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Pane;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utilerias.general.ControladorGeneral;

/**
 * FXML Controller class
 *
 * @author Samue
 */
public class PrincipalController implements Initializable {

    /**
     * Initializes the controller class.
     */
    @FXML
    private Button botonCerrar;
    @FXML
    private Button botonMinimizar;

    //Botones que redireccionan
    @FXML
    private Pane botonImportacion;

    @FXML
    private Pane botonVisualizacion;

    @FXML
    private Pane botonBusqueda;

    @FXML
    private Pane botonModificacion;

    @FXML
    private Pane botonExportacion;

    @FXML
    private Pane botonRespaldo;

    @FXML
    private ImageView notification;

    @FXML
    private Pane notificationPane;

    @FXML
    private Label cerrarNotificacion;

    //NUEVAS NOTIFICACIONES
    @FXML
    private VBox notificacioneBox;
    @FXML
    private ScrollPane scrollBox;

    @FXML
    private ImageView notiAlert;

    @FXML
    private ImageView reloadNotis;

    //Métodos de los botones de la barra superior :)
    public void cerrarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión?", getClass());
    }

    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }
    
      private boolean validarEstructuraArchivos() {
        // Obtener año y periodo actual
        Calendar calendario = Calendar.getInstance();
        int año = calendario.get(Calendar.YEAR);
        int periodo = (calendario.get(Calendar.MONTH) + 1) < 7 ? 1 : 2;

        // Rutas base para los directorios a validar
        String rutaBaseImportados = ControladorGeneral.obtenerRutaDeEjecusion()
                + File.separator + "Gestion_de_Cursos" 
                + File.separator + "Archivos_importados"
                + File.separator + año
                + File.separator + periodo + "-" + año;

        String rutaInfoModificable = ControladorGeneral.obtenerRutaDeEjecusion()
                + File.separator + "Gestion_de_Cursos"
                + File.separator + "Sistema"
                + File.separator + "informacion_modificable";

        // Subcarpetas a validar
        String[] subcarpetas = {
            "formato_de_hojas_membretadas_para_reconocimientos",
            "formato_de_lista_de_asistencias",
            "formato_de_reporte_para_docentes_capacitados",
            "listado_de_deteccion_de_necesidades",
            "listado_de_etiquetas_de_cursos",
            "listado_de_pre_regitro_a_cursos_de_capacitacion",
            "listado_de_docentes_adscritos",
            "programa_institucional"
        };
        
          System.out.println("DIR: "+rutaBaseImportados);
        // Validar estructura de subcarpetas de archivos importados
        File dirImportados = new File(rutaBaseImportados);
        if (!dirImportados.exists()) {
            deshabilitarBotones();
            return false;
        }

        // Verificar cada subcarpeta
        for (String subcarpeta : subcarpetas) {
            File dirSubcarpeta = new File(dirImportados, subcarpeta);
            if (!dirSubcarpeta.exists() || !tieneArchivos(dirSubcarpeta)) {
                deshabilitarBotones();
                return false;
            }
        }

        // Validar archivo info.xlsx
        File archivoInfo = new File(rutaInfoModificable, "info.xlsx");
        if (!archivoInfo.exists()) {
            deshabilitarBotones();
            return false;
        }

        // Validar hoja del año actual en info.xlsx
        try (FileInputStream fis = new FileInputStream(archivoInfo);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            boolean tieneHojaDelAño = false;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (workbook.getSheetName(i).equals(String.valueOf(año))) {
                    tieneHojaDelAño = true;
                    break;
                }
            }

            if (!tieneHojaDelAño) {
                deshabilitarBotones();
                return false;
            }
        } catch (IOException e) {
            deshabilitarBotones();
            return false;
        }

        return true;
    }

    // Método para verificar si un directorio tiene archivos
    private boolean tieneArchivos(File directorio) {
        File[] archivos = directorio.listFiles();
        return archivos != null && archivos.length > 0;
    }

     public void deshabilitarBotones(){
        botonBusqueda.setDisable(true);
        botonVisualizacion.setDisable(true);
        botonExportacion.setDisable(true);
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        /// Validar estructura de archivos antes de continuar
        validarEstructuraArchivos();

        // Asegurar que notificationPane esté cerrado inicialmente
        notificationPane.setVisible(false);
        notiAlert.setVisible(false);

        generarNotificacionesEnVBox();
        generarNotis();
        
         
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

        // VISTA UNO ()
        botonImportacion.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.regresar(event, "ImportacionArchivos", getClass());
            } catch (IOException ex) {
                Logger.getLogger(PrincipalController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        // VISTA DOS
        botonVisualizacion.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.regresar(event, "VizualizacionDatos", getClass());
            } catch (IOException ex) {
                Logger.getLogger(PrincipalController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        //VISTA TRES
        botonBusqueda.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.regresar(event, "BusquedaEstadistica", getClass());
            } catch (IOException ex) {
                Logger.getLogger(PrincipalController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        //VISTA CUATRO
        botonModificacion.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.regresar(event, "ModificacionDatos", getClass());
            } catch (IOException ex) {
                Logger.getLogger(PrincipalController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        //VISTA CINCO
        botonExportacion.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.regresar(event, "ExportacionReconocimientos", getClass());
            } catch (IOException ex) {
                Logger.getLogger(PrincipalController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        //VISTA SEIS
        botonRespaldo.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.regresar(event, "Respaldo", getClass());
            } catch (IOException ex) {
                Logger.getLogger(PrincipalController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        
        

        // Obtener año y periodo actual
        Calendar calendario = Calendar.getInstance();
        int año = calendario.get(Calendar.YEAR);
        int periodo = (calendario.get(Calendar.MONTH) + 1) < 7 ? 1 : 2;

        // Construir ruta base para archivo de salida
        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion()
                + File.separator + "Gestion_de_Cursos" + File.separator
                + "Sistema" + File.separator + "informacion_notificaciones"
                + File.separator + año + File.separator + periodo + "-" + año
                + File.separator;

        // Leer docentes que necesitan capacitación
        List<Docente> docentes = leerDocentesConNecesidadDeCapacitacion(rutaArchivo);

        // Configurar el botón de notificación para mostrar el `notificationPane`
        notification.setOnMouseClicked(event -> {
            notificationPane.setVisible(true);
        });

        // Configurar el botón para cerrar el `notificationPane`
        cerrarNotificacion.setOnMouseClicked(event -> {
            notificationPane.setVisible(false);
        });

        // Ocultar el panel inicialmente
        notificationPane.setVisible(false);

        reloadNotis.setOnMouseClicked(event -> {
            generarNotis();
            System.out.println("Generacion de archivo para notis");
        });
    }

    /////////////////////////////////////////////////////////////////////////////////////////////////
    //NOTIFICACIONES//
    /////////////////////////////////////////////////////////////////////////////////////////////////
    // Clase interna para almacenar la información de cada docente
    public static class Docente {

        String nombre;
        boolean necesitaCapacitacionFD;
        boolean necesitaCapacitacionAP;

        public Docente(String nombre, boolean necesitaFP, boolean necesitaAD) {
            this.nombre = nombre;
            this.necesitaCapacitacionFD = necesitaFP;
            this.necesitaCapacitacionAP = necesitaAD;
        }

    }

    //
    // Leer listado_de_necesidad_de_acreditacion con las columnas A, B y C
    public static List<Docente> leerDocentesConNecesidadDeCapacitacion(String rutaArchivo) {
        List<Docente> docentes = new ArrayList<>();

        // Verificar si el archivo existe antes de intentar leerlo
        File archivo = new File(rutaArchivo);
        if (!archivo.exists()) {
            System.out.println("El archivo especificado no existe: " + rutaArchivo); // Mensaje en consola
            // Alternativamente, puedes usar un sistema de notificaciones visual
            // o simplemente manejarlo con un log
            return docentes; // Retornar lista vacía si el archivo no existe
        }

        // Intentar leer el archivo si existe
        try (FileInputStream fis = new FileInputStream(archivo); Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // Leer la primera hoja

            // Iterar desde la segunda fila (índice 1) para omitir los encabezados
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue; // Saltar filas vacías
                }
                // Leer el nombre del docente en la columna A (índice 0)
                Cell nombreCell = row.getCell(0); // Columna A
                Cell fpCell = row.getCell(1);     // Columna B (FP)
                Cell adCell = row.getCell(2);     // Columna C (AD)

                // Validar que la celda del nombre no sea nula
                if (nombreCell != null) {
                    String nombre = nombreCell.toString().trim(); // Obtener valor como String

                    // Validar FP y AD, asignando "Recomendable" solo si coincide
                    boolean necesitaFP = fpCell != null && "Recomendable".equalsIgnoreCase(fpCell.toString().trim());
                    boolean necesitaAD = adCell != null && "Recomendable".equalsIgnoreCase(adCell.toString().trim());

                    // Agregar solo si se tiene un nombre válido
                    docentes.add(new Docente(nombre, necesitaFP, necesitaAD));
                }
            }
        } catch (Exception e) {
            // Manejar cualquier otro error que pueda ocurrir durante la lectura
            System.out.println("Error al leer el archivo: " + e.getMessage());
        }

        return docentes;
    }

    private String obtenerUltimaSemana(String rutaDirectorio, String patron, String identificador) {
        File directorio = new File(rutaDirectorio);
        System.out.println("Buscando archivos en: " + rutaDirectorio);

        if (!directorio.exists()) {
            System.out.println("ERROR: Directorio no existe: " + rutaDirectorio);
            return "0";
        }

        File[] archivos = directorio.listFiles();
        if (archivos == null || archivos.length == 0) {
            System.out.println("ERROR: No se encontraron archivos en el directorio");
            return "0";
        }

        System.out.println("Archivos encontrados: " + archivos.length);

        int ultimaSemana = 0;
        Pattern pattern = Pattern.compile(patron);

        for (File archivo : archivos) {
            System.out.println("Analizando archivo: " + archivo.getName());
            Matcher matcher = pattern.matcher(archivo.getName());
            if (matcher.find()) {
                try {
                    String numeroStr = archivo.getName().split(identificador + "_")[1].split("\\)")[0];
                    int numero = Integer.parseInt(numeroStr);
                    System.out.println("Número de semana encontrado: " + numero);
                    if (numero > ultimaSemana) {
                        ultimaSemana = numero;
                    }
                } catch (Exception e) {
                    System.out.println("Error al procesar archivo " + archivo.getName() + ": " + e.getMessage());
                }
            }
        }

        System.out.println("Última semana encontrada: " + ultimaSemana);
        return String.valueOf(ultimaSemana);
    }

    // Método para leer los datos de docentes y generar las notificaciones en VBox
    public void generarNotificacionesEnVBox() {
        // Obtener año y periodo actual
        Calendar calendario = Calendar.getInstance();
        int año = calendario.get(Calendar.YEAR);
        int periodo = (calendario.get(Calendar.MONTH) + 1) < 7 ? 1 : 2;

        // Construir ruta base para archivo de salida
        String rutaBaseSalida = ControladorGeneral.obtenerRutaDeEjecusion()
                + File.separator + "Gestion_de_Cursos" + File.separator
                + "Sistema" + File.separator + "informacion_notificaciones"
                + File.separator + año + File.separator + periodo + "-" + año
                + File.separator;

        System.out.println("Ruta base de salida: " + rutaBaseSalida);

        // Verificar si el directorio existe
        File directorioBase = new File(rutaBaseSalida);
        if (!directorioBase.exists()) {
            System.out.println("ERROR: El directorio base no existe: " + rutaBaseSalida);
            return;
        }

        // Obtener última semana
        String semanaCapacitacion = obtenerUltimaSemana(rutaBaseSalida,
                "docentes_recomendables_\\(Semana_\\d+\\)\\.xlsx", "Semana");

        System.out.println("Última semana encontrada: " + semanaCapacitacion);

        // Si no se encontró ninguna semana
        if (semanaCapacitacion.equals("0")) {
            System.out.println("ERROR: No se encontraron archivos de semanas");
            return;
        }

        // Construir ruta completa del archivo de salida
        String rutaSalida = rutaBaseSalida + "docentes_recomendables_(Semana_"
                + semanaCapacitacion + ").xlsx";

        System.out.println("Intentando leer archivo: " + rutaSalida);

        // Verificar si el archivo existe
        File archivoSalida = new File(rutaSalida);
        if (!archivoSalida.exists()) {
            System.out.println("ERROR: El archivo no existe: " + rutaSalida);
            return;
        }

        List<Docente> docentesN = leerDocentesConNecesidadDeCapacitacion(rutaSalida);
        System.out.println("Docentes leídos: " + docentesN.size());

        // Limpiar el VBox antes de agregar nuevas notificaciones
        notificacioneBox.getChildren().clear();

        int docentesMostrados = 0;
        // Crear una entrada de notificación para cada docente
        for (Docente docente : docentesN) {
            // Solo procesar docentes con recomendaciones
            if (!docente.necesitaCapacitacionFD && !docente.necesitaCapacitacionAP) {
                continue;
            }

            docentesMostrados++;

            // Crear un VBox para el docente
            VBox docenteBox = new VBox();
            docenteBox.setSpacing(3);
            docenteBox.setStyle("-fx-padding: 5; -fx-border-color: #cccccc; "
                    + "-fx-background-color: white; -fx-border-radius: 5; "
                    + "-fx-background-radius: 5; -fx-margin: 5;");

            Label nombreLabel = new Label("Nombre: " + docente.nombre);
            nombreLabel.setStyle("-fx-font-weight: bold;");

            String necesidades = "Necesita capacitación en: ";
            if (docente.necesitaCapacitacionFD && docente.necesitaCapacitacionAP) {
                necesidades += "FD y AP";
            } else if (docente.necesitaCapacitacionFD) {
                necesidades += "FD";
            } else if (docente.necesitaCapacitacionAP) {
                necesidades += "AP";
            }

            Label capacitacionLabel = new Label(necesidades);
            docenteBox.getChildren().addAll(nombreLabel, capacitacionLabel);
            notificacioneBox.getChildren().add(docenteBox);
        }

        System.out.println("Docentes mostrados en notificaciones: " + docentesMostrados);

        // Ajustar visibilidad del ícono de notificación
        notiAlert.setVisible(docentesMostrados > 0);

        // Configurar ScrollPane
        scrollBox.setContent(notificacioneBox);
        scrollBox.setFitToWidth(true);
        scrollBox.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        scrollBox.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
    }

    public void generarNotis() {
        try {
            Necesidadcapacitacion necesidadCapacitacion = new Necesidadcapacitacion();
            necesidadCapacitacion.generarArchivo();
            System.out.println("Generación de archivo para notificaciones completada.");
            generarNotificacionesEnVBox();
        } catch (IOException e) {
            System.err.println("Error al generar el archivo: " + e.getMessage());
        }
    }

}// FIN PrincipalController
