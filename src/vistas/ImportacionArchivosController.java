package vistas;

import java.io.File;
import static java.io.File.separator;
import java.io.IOException;
import java.net.URL;
import java.util.*;
import java.util.logging.*;
import javafx.collections.FXCollections;
import javafx.fxml.*;
import javafx.scene.control.*;
import javafx.scene.input.*;
import javafx.stage.*;
import utilerias.general.ControladorGeneral;

/**
 * FXML Controller class
 *
 * @author Samue
 */
public class ImportacionArchivosController implements Initializable {

    private File ultimoDirectorioUsado;

    private File programaCapacitacion;

    private Map<String, List<File>> listadosTemporales = new HashMap<>();
    private Map<String, List<File>> formatosTemporales = new HashMap<>();

    /**
     * Initializes the controller class.
     */
    @FXML
    private Button botonCerrar;
    @FXML
    private Button botonMinimizar;
    @FXML
    private Button botonRegresar;
    @FXML
    private Button botonPC;
    @FXML
    private Button botonListado;
    @FXML
    private Button botonFormato;
    @FXML
    private Button botonGuardar;
    @FXML
    private ComboBox<String> comboBoxListados;
    @FXML
    private ComboBox<String> comboBoxFormatos;
    @FXML
    private Label labelPC;
    @FXML
    private Label labelListado;
    @FXML
    private Label labelFormato;

    //Métodos de los botones de la barra superior :)
    public void cerrarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión?", getClass());
    }

    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    public void regresarVentana(MouseEvent event) throws IOException {
        // Verificar si hay archivos sin guardar
        boolean hayArchivosSinGuardar = (programaCapacitacion != null)
                || (!listadosTemporales.isEmpty())
                || (!formatosTemporales.isEmpty());

        if (hayArchivosSinGuardar) {
            // Crear un cuadro de diálogo de confirmación
            Alert confirmacion = new Alert(Alert.AlertType.CONFIRMATION);
            confirmacion.setTitle("Archivos sin guardar");
            confirmacion.setHeaderText("Tiene archivos sin guardar");
            confirmacion.setContentText("¿Desea guardar los cambios antes de salir?");

            // Añadir botones personalizados
            ButtonType botonGuardar = new ButtonType("Guardar");
            ButtonType botonSalirSinGuardar = new ButtonType("Salir sin guardar");
            ButtonType botonCancelar = new ButtonType("Cancelar", ButtonBar.ButtonData.CANCEL_CLOSE);

            confirmacion.getButtonTypes().setAll(botonGuardar, botonSalirSinGuardar, botonCancelar);

            // Mostrar el diálogo y obtener la respuesta
            Optional<ButtonType> resultado = confirmacion.showAndWait();

            if (resultado.get() == botonGuardar) {
                // Llamar al método de guardar archivos
                guardarArchivos(event);
                // Después de guardar, regresar a la ventana anterior
                ControladorGeneral.regresar(event, "Principal", getClass());
            } else if (resultado.get() == botonSalirSinGuardar) {
                // Salir sin guardar
                ControladorGeneral.regresar(event, "Principal", getClass());
            }
            // Si se selecciona Cancelar, no se hace nada (se queda en la ventana actual)
        } else {
            // Si no hay archivos sin guardar, simplemente regresar
            ControladorGeneral.regresar(event, "Principal", getClass());
        }
    }

    public void cargarProgramaCap(MouseEvent event) {
        FileChooser selectorArchivos = new FileChooser();
        selectorArchivos.setTitle("Seleccionar archivo");
        // Usar el último directorio usado o el directorio de usuario
        if (ultimoDirectorioUsado != null && ultimoDirectorioUsado.exists()) {
            selectorArchivos.setInitialDirectory(ultimoDirectorioUsado);
        } else {
            selectorArchivos.setInitialDirectory(new File(System.getProperty("user.home")));
        }

        selectorArchivos.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Archivos Excel", "*.xlsx", "*.xls")
        );

        programaCapacitacion = selectorArchivos.showOpenDialog(botonPC.getScene().getWindow());
        if (programaCapacitacion != null) {
            // Actualizar último directorio usado
            ultimoDirectorioUsado = programaCapacitacion.getParentFile();
            labelPC.setText("(1) archivo(s) cargados");
        }

    }

    public void cargarListado(MouseEvent event) {
        if (comboBoxListados.getValue() == null) {
            mostrarError("Debe seleccionar primero una opción de listado.");
            return;
        }

        FileChooser selectorArchivos = new FileChooser();
        selectorArchivos.setTitle("Seleccionar archivo");

        if (ultimoDirectorioUsado != null && ultimoDirectorioUsado.exists()) {
            selectorArchivos.setInitialDirectory(ultimoDirectorioUsado);
        } else {
            selectorArchivos.setInitialDirectory(new File(System.getProperty("user.home")));
        }

        selectorArchivos.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Archivos Excel", "*.xlsx", "*.xls")
        );

        File listadoCargado = selectorArchivos.showOpenDialog(botonPC.getScene().getWindow());
        if (listadoCargado != null) {
            ultimoDirectorioUsado = listadoCargado.getParentFile();

            String tipoListado = comboBoxListados.getValue();

            // Si ya existe este tipo de listado, reemplazamos el archivo
            if (listadosTemporales.containsKey(tipoListado)) {
                listadosTemporales.get(tipoListado).clear();
            } else {
                // Si no existe, creamos una nueva lista
                listadosTemporales.put(tipoListado, new ArrayList<>());
            }

            // Agregamos el nuevo archivo
            listadosTemporales.get(tipoListado).add(listadoCargado);

            // Calculamos el total de tipos de listados diferentes cargados
            int totalListados = listadosTemporales.size();

            // Actualizamos el label
            labelListado.setText(String.format("(%d) archivo(s) cargados", totalListados));
        }
    }

    public void cargarFormato(MouseEvent event) {
        if (comboBoxFormatos.getValue() == null) {
            mostrarError("Debe seleccionar primero una opción de formato.");
            return;
        }

        FileChooser selectorArchivos = new FileChooser();
        selectorArchivos.setTitle("Seleccionar archivo");

        if (ultimoDirectorioUsado != null && ultimoDirectorioUsado.exists()) {
            selectorArchivos.setInitialDirectory(ultimoDirectorioUsado);
        } else {
            selectorArchivos.setInitialDirectory(new File(System.getProperty("user.home")));
        }

        String tipoFormato = comboBoxFormatos.getValue();
        if (tipoFormato.equals("Formato de hojas membretadas para reconocimientos")) {
            selectorArchivos.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Archivos Word", "*.doc", "*.docx")
            );
        } else {
            selectorArchivos.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Archivos Excel", "*.xlsx", "*.xls")
            );
        }

        File formatoCargado = selectorArchivos.showOpenDialog(botonPC.getScene().getWindow());
        if (formatoCargado != null) {
            ultimoDirectorioUsado = formatoCargado.getParentFile();

            String tipoFormat = comboBoxFormatos.getValue();

            // Si ya existe este tipo de formato, reemplazamos el archivo
            if (formatosTemporales.containsKey(tipoFormat)) {
                formatosTemporales.get(tipoFormat).clear();
            } else {
                // Si no existe, creamos una nueva lista
                formatosTemporales.put(tipoFormat, new ArrayList<>());
            }

            // Agregamos el nuevo archivo
            formatosTemporales.get(tipoFormat).add(formatoCargado);

            // Calculamos el total de tipos de formatos diferentes cargados
            int totalFormatos = formatosTemporales.size();

            // Actualizamos el label
            labelFormato.setText(String.format("(%d) archivo(s) cargados", totalFormatos));
        }
    }

    private void mostrarError(String mensaje) {
        Alert alert = new Alert(Alert.AlertType.ERROR);
        alert.setTitle("Error");
        alert.setHeaderText(null);
        alert.setContentText(mensaje);
        alert.showAndWait();
    }

    private void inicializarEstructuraDirectorios() {
        String directorioUsuario = System.getProperty("user.home");
        String separador = File.separator;
        String directorioBase = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos";

        // Obtener año actual
        Calendar calendario = Calendar.getInstance();
        int year = calendario.get(Calendar.YEAR);
        int mesActual = calendario.get(Calendar.MONTH) + 1;

        // Definir la estructura de directorios
        String carpetaPeriodo = (mesActual >= 1 && mesActual <= 7) ? "1-" + year : "2-" + year;
        String directorioImportados = directorioBase + separador + "Archivos_importados" + separador + year + separador + carpetaPeriodo;
        String directorioExportados = directorioBase + separador + "Archivos_exportados" + separador + year + separador + carpetaPeriodo;
        String directorioSistema = directorioBase + separador + "Sistema";

        // Crear directorio base
        crearDirectorio(directorioBase);

        // Crear estructura de Importados y Exportados
        crearDirectorio(directorioImportados);
        crearDirectorio(directorioExportados);

        // Crear subdirectorios de archivos importados
        crearDirectorio(directorioImportados + separador + "formato_de_hojas_membretadas_para_reconocimientos");
        crearDirectorio(directorioImportados + separador + "formato_de_lista_de_asistencias");
        crearDirectorio(directorioImportados + separador + "formato_de_reporte_para_docentes_capacitados");
        crearDirectorio(directorioImportados + separador + "listado_de_pre_regitro_a_cursos_de_capacitacion");
        crearDirectorio(directorioImportados + separador + "listado_de_etiquetas_de_cursos");
        crearDirectorio(directorioImportados + separador + "listado_de_deteccion_de_necesidades");
        crearDirectorio(directorioImportados + separador + "listado_de_docentes_adscritos");
        crearDirectorio(directorioImportados + separador + "programa_institucional");

        // Crear subdirectorios de archivos exportados
        crearDirectorio(directorioExportados + separador + "listas_asistencia");
        crearDirectorio(directorioExportados + separador + "reconocimientos");
        crearDirectorio(directorioExportados + separador + "reportes_estadisticos");

        // Crear directorios de Sistema
        crearDirectorio(directorioSistema);

        // Crear 4 directorios principales en Sistema
        String condensadosVista = directorioSistema + separador + "condensados_vista_de_visualizacion_de_datos";
        String informacionModificable = directorioSistema + separador + "informacion_modificable";
        String informacionNotificaciones = directorioSistema + separador + "informacion_notificaciones";
        String registrosContrasenas = directorioSistema + separador + "registros_contrasenas";

        crearDirectorio(condensadosVista);
        crearDirectorio(informacionModificable);
        crearDirectorio(informacionNotificaciones);
        crearDirectorio(registrosContrasenas);

        // Crear subdirectorios para condensados_vista_de_visualizacion_de_datos y informacion_notificaciones
        crearDirectorio(condensadosVista + separador + year + separador + carpetaPeriodo);
        crearDirectorio(informacionNotificaciones + separador + year + separador + carpetaPeriodo);
    }

    private void crearDirectorio(String ruta) {
        File directorio = new File(ruta);
        if (!directorio.exists()) {
            directorio.mkdirs();
        }
    }

    public void guardarArchivos(MouseEvent evento) {
        String directorioBase = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos";
        String separador = File.separator;

        // Determinar el período actual
        Calendar calendario = Calendar.getInstance();
        int year = calendario.get(Calendar.YEAR);
        int mesActual = calendario.get(Calendar.MONTH) + 1;

        String carpetaPeriodo = (mesActual >= 1 && mesActual <= 7)
                ? "1-" + year
                : "2-" + year;

        String directorioImportados = directorioBase + separador + "Archivos_importados"
                + separador + year + separador + carpetaPeriodo;

        int totalArchivosImportados = 0;

        // Guardar Programa de Capacitación
        if (programaCapacitacion != null) {
            String dirPC = directorioImportados + separador + "programa_institucional";
            totalArchivosImportados += guardarArchivoEnDirectorio(programaCapacitacion, dirPC, "programa_institucional");
        }

        // Guardar Listados
        for (Map.Entry<String, List<File>> entry : listadosTemporales.entrySet()) {
            String tipoListado = entry.getKey();
            List<File> archivosListado = entry.getValue();
            String nombreCarpeta = convertirNombreCarpeta(tipoListado);
            String dirListado = directorioImportados + separador + nombreCarpeta;

            for (File archivo : archivosListado) {
                totalArchivosImportados += guardarArchivoEnDirectorio(archivo, dirListado, "listado");
            }
        }

        // Guardar Formatos
        for (Map.Entry<String, List<File>> entry : formatosTemporales.entrySet()) {
            String tipoFormato = entry.getKey();
            List<File> archivosFormato = entry.getValue();
            String nombreCarpeta = convertirNombreCarpeta(tipoFormato);
            String dirFormato = directorioImportados + separador + nombreCarpeta;

            for (File archivo : archivosFormato) {
                totalArchivosImportados += guardarArchivoEnDirectorio(archivo, dirFormato, "formato");
            }
        }

        // Mostrar mensaje de éxito y limpiar
        if (totalArchivosImportados > 0) {
            mostrarMensajeExito("Se importaron correctamente " + totalArchivosImportados + " archivo(s)");
            limpiarCampos();
        } else {
            mostrarError("No se ha seleccionado ningún archivo para importar");
        }
    }

    private int guardarArchivoEnDirectorio(File archivo, String directorioDestino, String prefijo) {
        if (archivo == null) {
            return 0;
        }

        File dirDestino = new File(directorioDestino);
        dirDestino.mkdirs();

        int numeroSemana = 1;
        String extension = getExtensionArchivo(archivo);
        File archivoDestino;
        do {
            // Usar Version para formatos y Semana para listados, con la primera letra en mayúscula
            String numeracion = prefijo.equals("formato")
                    ? "_(Version_" + numeroSemana + ")"
                    : "_(Semana_" + numeroSemana + ")";

            archivoDestino = new File(directorioDestino + separator
                    + prefijo + numeracion + extension);

            numeroSemana++;
        } while (archivoDestino.exists());

        try {
            java.nio.file.Files.copy(archivo.toPath(), archivoDestino.toPath());
            return 1;
        } catch (IOException e) {
            mostrarError("Error al copiar archivo: " + e.getMessage());
            return 0;
        }
    }

    private void limpiarCampos() {
        labelPC.setText("");
        labelListado.setText("");
        labelFormato.setText("");

        programaCapacitacion = null;
        listadosTemporales.clear();
        formatosTemporales.clear();

        comboBoxListados.setValue(null);
        comboBoxFormatos.setValue(null);
        comboBoxListados.setPromptText("Selecciona una opción");
        comboBoxFormatos.setPromptText("Selecciona una opción");

    }

// Método auxiliar para obtener la extensión del archivo
    private String getExtensionArchivo(File archivo) {
        String nombreArchivo = archivo.getName();
        int lastIndexOf = nombreArchivo.lastIndexOf(".");
        if (lastIndexOf == -1) {
            return "";
        }
        return nombreArchivo.substring(lastIndexOf);
    }

// Método auxiliar para convertir nombres de ComboBox a nombres de carpeta
    private String convertirNombreCarpeta(String nombre) {
        return nombre.toLowerCase()
                .replace(" ", "_")
                .replace("á", "a")
                .replace("é", "e")
                .replace("í", "i")
                .replace("ó", "o")
                .replace("ú", "u")
                .replace("ñ", "n");
    }

// Método para mostrar mensaje de éxito
    private void mostrarMensajeExito(String mensaje) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle("Éxito");
        alert.setHeaderText(null);
        alert.setContentText(mensaje);
        alert.showAndWait();
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {

        inicializarEstructuraDirectorios();

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

        comboBoxListados.setValue(null);
        comboBoxFormatos.setValue(null);

        comboBoxListados.setItems(FXCollections.observableArrayList(
                "Listado de pre regitro a cursos de capacitación",
                "Listado de etiquetas de cursos",
                "Listado de detección de necesidades",
                "Listado de docentes adscritos"));

        comboBoxFormatos.setItems(FXCollections.observableArrayList(
                "Formato de hojas membretadas para reconocimientos",
                "Formato de lista de asistencias",
                "Formato de reporte para docentes capacitados"));

        comboBoxListados.setPromptText("Selecciona una opción");
        comboBoxFormatos.setPromptText("Selecciona una opción");

        botonPC.setOnMouseClicked(event -> {
            cargarProgramaCap(event);
        });

        botonListado.setOnMouseClicked(event -> {
            cargarListado(event);
        });

        botonFormato.setOnMouseClicked(event -> {
            cargarFormato(event);
        });

        botonGuardar.setOnMouseClicked(event -> {
            guardarArchivos(event);
        });

    }

}