package vistas;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.net.URL;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Optional;
import java.util.ResourceBundle;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import javafx.application.Application;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.ListView;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.cell.CheckBoxTableCell;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.ImageView;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import utilerias.general.ControladorGeneral;
import org.apache.poi.ss.util.CellReference;

public class VizualizacionDatosController implements Initializable {

    @FXML
    private Button botonCerrar, botonMinimizar, botonRegresar, botonExportar,
            botonActualizar, botonModificar, botonGuardar, botonBuscar;

    @FXML
    private TextField campoBusqueda;

    @FXML
    private TableView<Evento> tableView;

    private ObservableList<Evento> eventoList = FXCollections.observableArrayList();

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        botonCerrar.setOnMouseClicked(event -> {
            try {
                cerrarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(VizualizacionDatosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        botonMinimizar.setOnMouseClicked(this::minimizarVentana);
        botonRegresar.setOnMouseClicked(event -> {
            try {
                regresarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(VizualizacionDatosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        campoBusqueda.setOnKeyReleased(event -> {
            Buscar();
        });

        configurarTabla();
        cargarDatos();
    }

    private void configurarTabla() {
        TableColumn<Evento, String> colApellidoPaterno = new TableColumn<>("Apellido Paterno");
        colApellidoPaterno.setCellValueFactory(new PropertyValueFactory<>("apellidoPaterno"));

        TableColumn<Evento, String> colApellidoMaterno = new TableColumn<>("Apellido Materno");
        colApellidoMaterno.setCellValueFactory(new PropertyValueFactory<>("apellidoMaterno"));

        TableColumn<Evento, String> colNombres = new TableColumn<>("Nombres");
        colNombres.setCellValueFactory(new PropertyValueFactory<>("nombres"));

        TableColumn<Evento, String> colRFC = new TableColumn<>("RFC");
        colRFC.setCellValueFactory(new PropertyValueFactory<>("rfc"));

        TableColumn<Evento, String> colSexo = new TableColumn<>("Sexo");
        colSexo.setCellValueFactory(new PropertyValueFactory<>("sexo"));

        TableColumn<Evento, String> colDepartamento = new TableColumn<>("Departamento");
        colDepartamento.setCellValueFactory(new PropertyValueFactory<>("departamento"));

        TableColumn<Evento, String> colPuesto = new TableColumn<>("Puesto");
        colPuesto.setCellValueFactory(new PropertyValueFactory<>("puesto"));

        TableColumn<Evento, String> colNombreEvento = new TableColumn<>("Nombre del Evento");
        colNombreEvento.setCellValueFactory(new PropertyValueFactory<>("nombreEvento"));

        TableColumn<Evento, String> colCapacitacion = new TableColumn<>("Capacitación");
        colCapacitacion.setCellValueFactory(new PropertyValueFactory<>("capacitacion"));

        TableColumn<Evento, String> colNombreFacilitador = new TableColumn<>("Nombre del Facilitador");
        colNombreFacilitador.setCellValueFactory(new PropertyValueFactory<>("nombreFacilitador"));

        TableColumn<Evento, String> colPeriodo = new TableColumn<>("Periodo");
        colPeriodo.setCellValueFactory(new PropertyValueFactory<>("periodo"));

        // Columna de botones
        TableColumn<Evento, Void> colAcreditacion = new TableColumn<>("Acreditación");
        colAcreditacion.setCellFactory(getButtonCellFactory());

        tableView.getColumns().addAll(colApellidoPaterno, colApellidoMaterno, colNombres, colRFC, colSexo, colDepartamento,
                colPuesto, colNombreEvento, colCapacitacion, colNombreFacilitador, colPeriodo, colAcreditacion);
    }

    private Callback<TableColumn<Evento, Void>, TableCell<Evento, Void>> getButtonCellFactory() {
        return param -> new TableCell<>() {
            private final Button btnAcredita = new Button();
            private final Button btnNoAcredita = new Button();

            {
                // Cargar las imágenes desde los recursos
                Image imagePalomita = new Image(getClass().getResourceAsStream("/utilerias/visualizacionDatos/Palomita.png"));
                Image imageTacha = new Image(getClass().getResourceAsStream("/utilerias/visualizacionDatos/Tacha.png"));

                // Crear ImageView para redimensionar las imágenes
                ImageView iconPalomita = new ImageView(imagePalomita);
                iconPalomita.setFitWidth(16); // Ancho del ícono
                iconPalomita.setFitHeight(16); // Alto del ícono

                ImageView iconTacha = new ImageView(imageTacha);
                iconTacha.setFitWidth(16); // Ancho del ícono
                iconTacha.setFitHeight(16); // Alto del ícono

                // Asignar los íconos redimensionados a los botones
                btnAcredita.setGraphic(iconPalomita);
                btnNoAcredita.setGraphic(iconTacha);

                // Configurar acciones de los botones
                btnAcredita.setOnAction(event -> {
                    Evento evento = getTableView().getItems().get(getIndex());
                    evento.setAcreditado(true);
                    updateButtons(evento); // Actualiza la interfaz
                });

                btnNoAcredita.setOnAction(event -> {
                    Evento evento = getTableView().getItems().get(getIndex());
                    evento.setAcreditado(false);
                    updateButtons(evento); // Actualiza la interfaz
                });
            }

            private void updateButtons(Evento evento) {
                // Cambia los gráficos de los botones según el estado
                if (evento.isAcreditado()) {
                    ImageView iconPalomita = new ImageView(new Image(getClass().getResourceAsStream("/utilerias/visualizacionDatos/Palomita.png")));
                    iconPalomita.setFitWidth(16);
                    iconPalomita.setFitHeight(16);
                    btnAcredita.setGraphic(iconPalomita);
                    btnNoAcredita.setGraphic(null); // Ocultar el botón de tacha
                } else {
                    ImageView iconTacha = new ImageView(new Image(getClass().getResourceAsStream("/utilerias/visualizacionDatos/Tacha.png")));
                    iconTacha.setFitWidth(16);
                    iconTacha.setFitHeight(16);
                    btnNoAcredita.setGraphic(iconTacha);
                    btnAcredita.setGraphic(null); // Ocultar el botón de palomita
                }
            }

            @Override
            protected void updateItem(Void item, boolean empty) {
                super.updateItem(item, empty);
                if (empty) {
                    setGraphic(null);
                } else {
                    Evento evento = getTableView().getItems().get(getIndex());
                    updateButtons(evento);

                    // Organizar los botones en un VBox para alineación vertical
                    VBox vbox = new VBox(btnAcredita, btnNoAcredita);
                    vbox.setSpacing(5);
                    setGraphic(vbox);
                }
            }
        };
    }

    private void cargarDatos() {
        try {
            // Definir las rutas para ambos archivos Excel
            Calendar calendario = Calendar.getInstance();
            int año = calendario.get(Calendar.YEAR);
            int periodo = (calendario.get(Calendar.MONTH) + 1) < 7 ? 1 : 2;

            String rutaArchivoEventos = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_importados\\" + año + "\\" + periodo + "-" + año + "\\listado_de_pre_regitros_a_cursos_de_capacitacion\\";
            int numeroSemana = obtenerUltimaSemana(rutaArchivoEventos, "listado\\_\\(Semana_\\d+\\)\\.xlsx", "Semana");
            rutaArchivoEventos += "listado_(Semana_" + numeroSemana + ").xlsx";
            System.out.println(rutaArchivoEventos);

            String rutaArchivoClasificaciones = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Archivos_importados\\" + año + "\\" + periodo + "-" + año + "\\programa_institucional\\"; // Asegúrate de poner la ruta correcta de tu archivo de clasificaciones
            numeroSemana = obtenerUltimaSemana(rutaArchivoClasificaciones, "programa_institucional\\_\\(Semana_\\d+\\)\\.xlsx", "Semana");
            rutaArchivoClasificaciones += "programa_institucional_(Semana_" + numeroSemana + ").xlsx";
            System.out.println(rutaArchivoClasificaciones);

            // Llamar al método para leer los eventos y las clasificaciones
            eventoList.addAll(leerEventosDesdeExcel(rutaArchivoEventos, rutaArchivoClasificaciones));

            // Establecer los datos en la tabla
            tableView.setItems(eventoList);

        } catch (IOException e) {
            // Mostrar un mensaje de error si no se pudieron cargar los datos
            mostrarAlerta("Error", "No se pudieron cargar los datos del archivo Excel.", AlertType.ERROR);
        }
    }

    public int obtenerUltimaSemana(String carpetaDestino, String nombreArchivo, String versionS) {
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
        Pattern pattern = Pattern.compile(".*\\(" + versionS + "_(\\d+)\\)\\.xlsx$"); // Patrón para extraer el número de versión

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

// Método auxiliar para limpiar texto
    private String limpiarTextoInicial(String texto) {
        if (texto == null || texto.isEmpty()) {
            return "";
        }
        // Elimina el patrón "número." del inicio del texto y elimina espacios extras
        return texto.replaceFirst("^\\d+\\.s*", "").trim();
    }

    @FXML
    private void Exportar(ActionEvent event) {
        // Preguntar si desea exportar un solo curso o todos los cursos
        Alert opcionExportacion = new Alert(AlertType.CONFIRMATION);
        opcionExportacion.setTitle("Opciones de Exportación");
        opcionExportacion.setHeaderText("Seleccione el tipo de exportación");
        opcionExportacion.setContentText("¿Desea exportar la lista de asistencia de un curso específico o de todos los cursos?");

        ButtonType botonUnCurso = new ButtonType("Un Curso");
        ButtonType botonTodosLosCursos = new ButtonType("Todos los Cursos");
        ButtonType botonCancelar = new ButtonType("Cancelar", ButtonBar.ButtonData.CANCEL_CLOSE);

        opcionExportacion.getButtonTypes().setAll(botonUnCurso, botonTodosLosCursos, botonCancelar);
        Optional<ButtonType> opcionSeleccionada = opcionExportacion.showAndWait();

        if (opcionSeleccionada.isEmpty() || opcionSeleccionada.get() == botonCancelar) {
            return;
        }

        if (opcionSeleccionada.get() == botonUnCurso) {
            // Mostrar CheckList para seleccionar un curso
            Dialog<String> dialog = new Dialog<>();
            dialog.setTitle("Seleccionar Curso");
            dialog.setHeaderText("Seleccione un curso de la lista");

            ListView<String> listView = new ListView<>();

            // Eliminar duplicados utilizando distinct() en el stream
            listView.getItems().addAll(
                    eventoList.stream()
                            .map(Evento::getNombreEvento)
                            .distinct() // Esto elimina los duplicados
                            .collect(Collectors.toList())
            );

            dialog.getDialogPane().setContent(listView);
            ButtonType botonSeleccionar = new ButtonType("Seleccionar", ButtonBar.ButtonData.OK_DONE);
            dialog.getDialogPane().getButtonTypes().addAll(botonSeleccionar, ButtonType.CANCEL);

            dialog.setResultConverter(dialogButton -> {
                if (dialogButton == botonSeleccionar) {
                    return listView.getSelectionModel().getSelectedItem();
                }
                return null;
            });

            Optional<String> resultado = dialog.showAndWait();
            if (!resultado.isPresent()) {
                return;
            }

            String cursoIngresado = limpiarTextoInicial(resultado.get().trim());

            if (cursoIngresado.isEmpty()) {
                mostrarAlerta("No encontrado", "No se encontraron registros para el curso especificado.", AlertType.WARNING);
                return;
            }

            // Buscar eventos que coincidan con el curso (ignorando números iniciales)
            List<Evento> eventosFiltrados = eventoList.stream()
                    .filter(e -> limpiarTextoInicial(e.getNombreEvento())
                    .toLowerCase()
                    .contains(cursoIngresado.toLowerCase()))
                    .collect(Collectors.toList());

            if (eventosFiltrados.isEmpty()) {
                mostrarAlerta("No encontrado", "No se encontraron registros para el curso especificado.", AlertType.WARNING);
                return;
            }

            // Exportar el curso seleccionado
            try {
                exportarCurso(eventosFiltrados, cursoIngresado);
                mostrarAlerta("Exportación exitosa",
                        "Se exportó correctamente la lista de asistencia del curso: " + cursoIngresado,
                        AlertType.INFORMATION);
            } catch (Exception e) {
                mostrarAlerta("Error", "No se pudo exportar la lista de asistencia del curso: " + cursoIngresado, AlertType.ERROR);
            }

        } else if (opcionSeleccionada.get() == botonTodosLosCursos) {
            // Seguimiento de cursos exportados
            int cursosExportados = 0;
            int cursosConError = 0;
            Set<String> cursosProcesados = new HashSet<>();

            for (Evento evento : eventoList) {
                String cursoIngresado = limpiarTextoInicial(evento.getNombreEvento().trim());

                // Evitar procesar cursos duplicados
                if (cursosProcesados.contains(cursoIngresado)) {
                    continue;
                }

                List<Evento> eventosFiltrados = eventoList.stream()
                        .filter(e -> limpiarTextoInicial(e.getNombreEvento())
                        .toLowerCase()
                        .contains(cursoIngresado.toLowerCase()))
                        .collect(Collectors.toList());

                if (!eventosFiltrados.isEmpty()) {
                    try {
                        exportarCurso(eventosFiltrados, cursoIngresado);
                        cursosExportados++;
                        cursosProcesados.add(cursoIngresado);
                    } catch (Exception e) {
                        cursosConError++;
                        System.err.println("Error exportando curso: " + cursoIngresado);
                    }
                }
            }

            // Mostrar mensaje de resumen
            String mensajeResumen = "Se exportaron exitosamente " + cursosExportados + " listas de asistencia ";
            if (cursosConError > 0) {
                mensajeResumen += " (" + cursosConError + " cursos no se pudieron exportar)";
            }

            mostrarAlerta("Exportación exitosa!",
                    mensajeResumen,
                    AlertType.INFORMATION);
        }
    }

    private void exportarCurso(List<Evento> eventosFiltrados, String cursoIngresado) throws InvalidFormatException {
        try {
            // Configuración de rutas
            String gestionCursosPath = ControladorGeneral.obtenerRutaDeEjecusion() + File.separator + "Gestion_de_Cursos";

            LocalDate fechaActual = LocalDate.now();
            int mesActual = fechaActual.getMonthValue();
            int añoActual = fechaActual.getYear();
            String periodo = mesActual >= 1 && mesActual <= 7 ? "1" : "2";

            String carpetaPeriodo = periodo + "-" + añoActual;
            String rutaCarpetaFormatos = gestionCursosPath + File.separator + "Archivos_importados"
                    + File.separator + añoActual + File.separator + carpetaPeriodo
                    + File.separator + "formato_de_lista_de_asistencia";

            // Buscar última versión del formato
            File carpetaFormatos = new File(rutaCarpetaFormatos);
            if (!carpetaFormatos.exists() || !carpetaFormatos.isDirectory()) {
                throw new IOException("No se encontró la carpeta de formatos en: " + rutaCarpetaFormatos);
            }

            File[] archivosFormato = carpetaFormatos.listFiles((dir, name)
                    -> name.toLowerCase().matches("formato_\\(version_\\d+\\)\\.xlsx")
            );

            if (archivosFormato == null || archivosFormato.length == 0) {
                throw new IOException("No se encontraron formatos en: " + rutaCarpetaFormatos);
            }

            File ultimaVersion = Arrays.stream(archivosFormato)
                    .max((f1, f2) -> {
                        int v1 = extraerNumeroVersion(f1.getName());
                        int v2 = extraerNumeroVersion(f2.getName());
                        return Integer.compare(v1, v2);
                    })
                    .orElseThrow(() -> new IOException("No se pudo determinar la última versión del formato"));

            // Cargar plantilla y obtener la hoja
            Workbook workbook = null;
            try {
                workbook = WorkbookFactory.create(new FileInputStream(ultimaVersion));
            } catch (InvalidFormatException | EncryptedDocumentException ex) {
                Logger.getLogger(VizualizacionDatosController.class.getName()).log(Level.SEVERE, null, ex);
                throw new IOException("Error al abrir el formato: " + ex.getMessage());
            }
            Sheet sheet = workbook.getSheetAt(0);

            // Obtener el primer evento y trabajar con los nombres limpios
            Evento primerEvento = eventosFiltrados.get(0);

            // Llenar el formato usando los nombres limpios
            llenarFormato(sheet, cursoIngresado, primerEvento);

            // Crear estilo para las celdas de la tabla
            CellStyle style = crearEstiloCelda(workbook);

            // Llenar la tabla de participantes
            llenarTablaParticipantes(sheet, eventosFiltrados, style);

            // Crear directorios para exportación
            String rutaExportacion = gestionCursosPath + File.separator + "Archivos_exportados"
                    + File.separator + añoActual + File.separator + carpetaPeriodo
                    + File.separator + "listas_asistencia";
            new File(rutaExportacion).mkdirs();

            // Guardar el archivo
            String nombreArchivo = "lista_asistencia_" + cursoIngresado.replace(" ", "_") + ".xlsx";
            String rutaCompleta = rutaExportacion + File.separator + nombreArchivo;

            try (FileOutputStream fileOut = new FileOutputStream(rutaCompleta)) {
                workbook.write(fileOut);
                // No mostrar alerta individual aquí
            }

        } catch (IOException e) {
            // Registro de error para depuración
            e.printStackTrace();
            // No mostrar alerta individual
            throw new RuntimeException("Error al exportar curso: " + cursoIngresado, e);
        }
    }

    // Método para crear el estilo de celda
    private CellStyle crearEstiloCelda(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        return style;
    }

// Método principal que maneja el llenado del formato
    private void llenarFormato(Sheet sheet, String cursoBuscado, Evento primerEvento) throws InvalidFormatException {
        try {
            // Obtener información del curso del programa institucional
            InfoCurso infoCurso = obtenerInfoCurso(cursoBuscado);
            String nombreCoordinador = obtenerNombreCoordinador();

            // Crear estilo simple para las celdas
            // Llenar datos del encabezado
            setCellValue(sheet, "D8", cursoBuscado);                    // NOMBRE DEL EVENTO
            setCellValue(sheet, "D9", limpiarTextoInicial(primerEvento.getNombreFacilitador())); // NOMBRE DEL FACILITADOR
            setCellValue(sheet, "D10", primerEvento.getPeriodo());          // PERIODO
            setCellValue(sheet, "D11", "Instituto Tecnológico de Zacatepec");                       // SEDE

            setCellValue(sheet, "D38", limpiarTextoInicial(primerEvento.getNombreFacilitador()));
            // Establecer horario y duración obtenidos del programa institucional
            setCellValue(sheet, "Q9", infoCurso.horario);              // HORARIO
            setCellValue(sheet, "J9", String.valueOf(infoCurso.duracion) + " horas"); // DURACIÓN

            setCellValue(sheet, "I38", nombreCoordinador);
            // Tipo de capacitación
            String tipoCapacitacion = primerEvento.getCapacitacion().equals("AP")
                    ? "Acreditación Profesional"
                    : "Formación Docente";
            setCellValue(sheet, "K10", tipoCapacitacion);              // TIPO

        } catch (IOException e) {
            System.out.println(e);
            mostrarAlerta("Error", "Error al obtener información del programa institucional: " + e.getMessage(), AlertType.ERROR);
        }
    }

// Método helper para establecer valores en celdas específicas
    private void setCellValue(Sheet sheet, String cellReference, String value) {
        try {
            CellReference ref = new CellReference(cellReference);
            Row row = sheet.getRow(ref.getRow());
            if (row == null) {
                row = sheet.createRow(ref.getRow());
            }

            Cell cell = row.getCell(ref.getCol());
            if (cell == null) {
                cell = row.createCell(ref.getCol());
            }

            cell.setCellValue(value);
        } catch (Exception e) {
            System.err.println("Error al establecer valor en celda " + cellReference + ": " + e.getMessage());
        }
    }

    private String obtenerNombreCoordinador() throws IOException, InvalidFormatException {
        // Obtener la ruta base de ejecución
        String rutaBase = ControladorGeneral.obtenerRutaDeEjecusion()
                + File.separator + "Gestion_de_Cursos"
                + File.separator + "Sistema"
                + File.separator + "informacion_modificable"
                + File.separator + "info.xlsx";

        try (Workbook workbook = WorkbookFactory.create(new File(rutaBase))) {
            // Obtener el año actual
            int añoActual = LocalDate.now().getYear();

            // Buscar la hoja con el nombre del año actual
            Sheet sheet = null;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (workbook.getSheetName(i).equals(String.valueOf(añoActual))) {
                    sheet = workbook.getSheetAt(i);
                    break;
                }
            }

            // Si no se encuentra la hoja del año actual, lanzar una excepción
            if (sheet == null) {
                throw new IOException("No se encontró la hoja para el año " + añoActual);
            }

            // Obtener el valor de la celda B3 (coordinador)
            Row row = sheet.getRow(2); // Fila 3 (índice 2)
            Cell cell = row.getCell(1); // Columna B (índice 1)

            // Devolver el valor de la celda como cadena, o una cadena vacía si está vacía
            return cell != null ? cell.getStringCellValue().trim() : "";
        }
    }
// Método para llenar la tabla de participantes

    private void llenarTablaParticipantes(Sheet sheet, List<Evento> eventos, CellStyle style) {
        int rowNum = 15; // Comienza en la fila 15 (donde inicia la tabla de participantes)

        for (Evento evento : eventos) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                row = sheet.createRow(rowNum);
            }

            // Columnas según el formato mostrado
            createCell(row, 0, (rowNum - 14), style);  // No.
            createCell(row, 1, evento.getApellidoPaterno() + " "
                    + evento.getApellidoMaterno() + " "
                    + evento.getNombres(), style);  // NOMBRE DEL PARTICIPANTE
            createCell(row, 4, evento.getRfc(), style);   // R.F.C
            createCell(row, 5, evento.getDepartamento(), style);  // PUESTO Y DEPARTAMENTO DE ADSCRIPCIÓN

            // Género (H/M)
            if ("Hombre".equalsIgnoreCase(evento.getSexo())) {
                createCell(row, 9, "X", style);  // H
                createCell(row, 10, "", style);   // M
            } else if ("Mujer".equalsIgnoreCase(evento.getSexo())) {
                createCell(row, 9, "", style);   // H
                createCell(row, 10, "X", style);  // M
            }

            // Tipo de puesto (Base/Interinato)
            String puesto = evento.getPuesto().toUpperCase();
            if (puesto.contains("BASE")) {
                createCell(row, 11, "X", style);  // B
                createCell(row, 12, "", style);   // I
            } else if (puesto.contains("INTERINATO")) {
                createCell(row, 11, "", style);   // B
                createCell(row, 12, "X", style);  // I
            }

            rowNum++;
        }
    }

    // Método para obtener el archivo de programa institucional más reciente
    private File obtenerProgramaInstitucionalReciente(String año, String periodo) throws IOException {
        String rutaProgramas = ControladorGeneral.obtenerRutaDeEjecusion()
                + File.separator + "Gestion_de_Cursos"
                + File.separator + "Archivos_importados"
                + File.separator + año
                + File.separator + periodo
                + File.separator + "programa_institucional";

        File carpetaProgramas = new File(rutaProgramas);
        if (!carpetaProgramas.exists() || !carpetaProgramas.isDirectory()) {
            throw new IOException("No se encontró la carpeta de programas institucionales en: " + rutaProgramas);
        }

        File[] archivos = carpetaProgramas.listFiles((dir, name)
                -> name.toLowerCase().matches("programa_institucional_\\(semana_\\d+\\)\\.xlsx")
        );

        if (archivos == null || archivos.length == 0) {
            throw new IOException("No se encontraron archivos de programa institucional en: " + rutaProgramas);
        }

        return Arrays.stream(archivos)
                .max((f1, f2) -> {
                    int semana1 = extraerNumeroSemana(f1.getName());
                    int semana2 = extraerNumeroSemana(f2.getName());
                    return Integer.compare(semana1, semana2);
                })
                .orElseThrow(() -> new IOException("No se pudo determinar el archivo más reciente"));
    }

// Método para extraer el número de semana del nombre del archivo
    private int extraerNumeroSemana(String nombreArchivo) {
        Pattern pattern = Pattern.compile("semana_(\\d+)");
        Matcher matcher = pattern.matcher(nombreArchivo.toLowerCase());
        if (matcher.find()) {
            return Integer.parseInt(matcher.group(1));
        }
        return 0;
    }

// Clase para almacenar la información del curso
    private class InfoCurso {

        String horario;
        int duracion;

        public InfoCurso(String horario, int duracion) {
            this.horario = horario;
            this.duracion = duracion;
        }
    }

// Método para obtener horario y duración del curso
    private InfoCurso obtenerInfoCurso(String nombreEvento) throws IOException {
        LocalDate fechaActual = LocalDate.now();
        int mesActual = fechaActual.getMonthValue();
        int añoActual = fechaActual.getYear();
        String periodo = (mesActual >= 1 && mesActual <= 7) ? "1-" + añoActual : "2-" + añoActual;

        File archivoProgramaReciente = obtenerProgramaInstitucionalReciente(String.valueOf(añoActual), periodo);

        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(archivoProgramaReciente))) {
            Sheet sheet = workbook.getSheetAt(0);

            // Buscar el curso en el archivo
            for (Row row : sheet) {
                Cell cellNombreEvento = row.getCell(1); // Columna B: Nombre de los evento
                if (cellNombreEvento != null
                        && cellNombreEvento.getStringCellValue().trim().equalsIgnoreCase(nombreEvento.trim())) {

                    // Extraer horario del periodo de realización
                    Cell cellPeriodo = row.getCell(4); // Columna E: Periodo de Realización
                    String periodoCompleto = cellPeriodo.getStringCellValue();
                    String horario = extraerHorario(periodoCompleto);

                    // Extraer duración
                    Cell cellDuracion = row.getCell(6); // Columna G: No. de horas x Curso
                    int duracion = (int) cellDuracion.getNumericCellValue();

                    return new InfoCurso(horario, duracion);
                }
            }
            throw new IOException("No se encontró el curso especificado en el programa institucional");
        } catch (Exception e) {
            throw new IOException("Error al leer el archivo de programa institucional: " + e.getMessage());
        }
    }

// Método para extraer solo el horario de la cadena completa
    private String extraerHorario(String periodoCompleto) {
        Pattern pattern = Pattern.compile("(\\d{2}:\\d{2} a \\d{2}:\\d{2})");
        Matcher matcher = pattern.matcher(periodoCompleto);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return "";
    }

    // Método helper para extraer el número de versión del nombre del archivo
    private int extraerNumeroVersion(String nombreArchivo) {
        Pattern pattern = Pattern.compile("formato_\\(version_(\\d+)\\)\\.xlsx", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(nombreArchivo);
        if (matcher.find()) {
            return Integer.parseInt(matcher.group(1));
        }
        return 0; // Retorna 0 si no encuentra un número de versión
    }

    // Método helper para crear celdas de manera segura
    private void createCell(Row row, int column, Object value, CellStyle style) {
        Cell cell = row.createCell(column);
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        }
        cell.setCellStyle(style);
    }

// Método helper para establecer valores en celdas de manera segura
    private void setCellValueSafely(Sheet sheet, int rowIndex, int colIndex, String value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
    }

    @FXML
    private void mostrarAlerta(String titulo, String mensaje, AlertType tipo) {
        Alert alert = new Alert(tipo);
        alert.setTitle(titulo);
        alert.setContentText(mensaje);
        alert.showAndWait();
    }

    @FXML
    private void buscarPorEnter(KeyEvent event) {
        // Verifica si la tecla presionada es "Enter"
        if (event.getCode() == KeyCode.ENTER) {
            System.out.println("Enter presionado");
            Buscar(); // Llama al método de búsqueda
        }
    }

    @FXML
    private void Buscar() {
        String textoBusqueda = campoBusqueda.getText().trim().toLowerCase();

        // Si el campo de búsqueda está vacío, mostramos una advertencia.
        if (textoBusqueda.isEmpty()) {
            tableView.setItems(eventoList);
            return;
        }

        // Filtra los datos según el texto de búsqueda.
        FilteredList<Evento> datosFiltrados = new FilteredList<>(eventoList, p -> true);
        datosFiltrados.setPredicate(info -> {
            return info.getNombres().toLowerCase().contains(textoBusqueda)
                    || info.getApellidoPaterno().toLowerCase().contains(textoBusqueda)
                    || info.getApellidoMaterno().toLowerCase().contains(textoBusqueda)
                    || info.getRfc().toLowerCase().contains(textoBusqueda)
                    || info.getSexo().toLowerCase().contains(textoBusqueda)
                    || info.getNombreEvento().toLowerCase().contains(textoBusqueda)
                    || info.getDepartamento().toLowerCase().contains(textoBusqueda)
                    || info.getNombreFacilitador().toLowerCase().contains(textoBusqueda)
                    || info.getPeriodo().toLowerCase().contains(textoBusqueda)
                    || info.getPuesto().toLowerCase().contains(textoBusqueda);
        });

        // Si no se encuentra ningún resultado, mostramos un mensaje de información y restauramos la tabla completa.
        if (datosFiltrados.isEmpty()) {

            // Restauramos la tabla con todos los datos
            tableView.setItems(null);
        } else {
            // Si hay resultados, mostramos solo los datos filtrados.
            tableView.setItems(datosFiltrados);
        }

    }

    @FXML
    private void actuzalizar(ActionEvent event) {
        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("ImportacionArchivos.fxml"));
            Parent root = loader.load();

            Stage stage = (Stage) botonActualizar.getScene().getWindow(); // Obtener la ventana actual
            stage.setScene(new Scene(root)); // Mostrar la nueva escena
            stage.show();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @FXML
    private void modificar(ActionEvent event) {
        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("ModificacionDatos.fxml"));
            Parent root = loader.load();

            Stage stage = (Stage) botonModificar.getScene().getWindow();
            stage.setScene(new Scene(root));
            stage.show();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @FXML
    private void guardar(ActionEvent event) {
        // Lógica para guardar los datos
        tableView.setItems(eventoList);
        boolean guardadoExitoso = guardarDatos();

        // Mostrar mensaje de confirmación si se guarda correctamente
        if (guardadoExitoso) {
            Alert alert = new Alert(AlertType.INFORMATION);
            alert.setTitle("Confirmación de Guardado");
            alert.setHeaderText(null);
            alert.setContentText("Los datos han sido guardados correctamente.");
            alert.showAndWait();
        } else {
            // Mostrar mensaje de error en caso de falla
            Alert alert = new Alert(AlertType.ERROR);
            alert.setTitle("Error al Guardar");
            alert.setHeaderText(null);
            alert.setContentText("Hubo un error al guardar los datos.");
            alert.showAndWait();
        }
    }

    private boolean guardarDatos() {
        try {
            // Crear un libro de Excel
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Datos");

            // Crear una fila de encabezados
            Row headerRow = sheet.createRow(0);
            String[] columnas = {
                "ApellidoPaterno", "ApellidoMaterno", "Nombres",
                "RFC", "Sexo", "Departamento", "Puesto", "NombreEvento", "Capacitacion",
                "NombreFacilitador", "Periodo", "Acreditación"
            };

            // Añadir encabezados a la fila
            for (int i = 0; i < columnas.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columnas[i]);
            }

            // Añadir datos desde el TableView
            int rowIndex = 1; // Empezar en la segunda fila
            for (Evento evento : tableView.getItems()) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(evento.getApellidoPaterno());
                row.createCell(1).setCellValue(evento.getApellidoMaterno());
                row.createCell(2).setCellValue(evento.getNombres());
                row.createCell(3).setCellValue(evento.getRfc());
                row.createCell(4).setCellValue(evento.getSexo());
                row.createCell(5).setCellValue(evento.getDepartamento());
                row.createCell(6).setCellValue(evento.getPuesto());
                row.createCell(7).setCellValue(evento.getNombreEvento());
                row.createCell(8).setCellValue(evento.getCapacitacion());
                row.createCell(9).setCellValue(evento.getNombreFacilitador());
                row.createCell(10).setCellValue(evento.getPeriodo());
                row.createCell(11).setCellValue(evento.isAcreditado() ? "si" : "no");
            }

            // Autoajustar el ancho de las columnas
            for (int i = 0; i < columnas.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Guardar el archivo en C:\excel\datos_guardados.xlsx
            Calendar calendario = Calendar.getInstance();
            int año = calendario.get(Calendar.YEAR);
            int periodo = calendario.get(Calendar.MONTH) < 7 ? 1 : 2;

            String filePath = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos\\" + año + "\\" + periodo + "-" + año + "\\";
            int numeroSemana = obtenerUltimaSemana(filePath, "condensado\\_\\(version_\\d+\\)\\.xlsx", "version");
            filePath += "condensado_(version_" + (numeroSemana + 1) + ").xlsx";
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }

            workbook.close();
            return true; // Indica que se guardó correctamente

        } catch (IOException e) {
            e.printStackTrace(); // Imprime el error para depuración
            return false; // Indica que hubo un error al guardar
        }
    }

    //Métodos de los botones de la barra superior :)
    public void cerrarVentana(MouseEvent event) throws IOException {
        ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión?", getClass());
    }

    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    @FXML
    public void regresarVentana(MouseEvent event) throws IOException {
        // Check if there are any unsaved changes
        if (hayCambiosSinGuardar()) {
            // Mostrar diálogo de confirmación
            Alert confirmacion = new Alert(Alert.AlertType.CONFIRMATION);
            confirmacion.setTitle("Cambios sin guardar");
            confirmacion.setHeaderText("Hay cambios sin guardar");
            confirmacion.setContentText("¿Desea guardar los cambios antes de salir?");

            // Añadir botones personalizados
            ButtonType btnGuardar = new ButtonType("Guardar");
            ButtonType btnSalirSinGuardar = new ButtonType("Salir sin guardar");
            ButtonType btnCancelar = new ButtonType("Cancelar", ButtonBar.ButtonData.CANCEL_CLOSE);
            confirmacion.getButtonTypes().setAll(btnGuardar, btnSalirSinGuardar, btnCancelar);

            // Mostrar y esperar la respuesta del usuario
            Optional<ButtonType> resultado = confirmacion.showAndWait();

            if (resultado.get() == btnGuardar) {
                // Intentar guardar los datos
                boolean guardadoExitoso = guardarDatos();
                if (guardadoExitoso) {
                    // Si se guardó correctamente, regresar a la ventana anterior
                    ControladorGeneral.regresar(event, "Principal", getClass());
                } else {
                    // Si hubo un error al guardar, mostrar mensaje de error
                    Alert errorAlGuardar = new Alert(Alert.AlertType.ERROR);
                    errorAlGuardar.setTitle("Error");
                    errorAlGuardar.setHeaderText(null);
                    errorAlGuardar.setContentText("No se pudieron guardar los cambios.");
                    errorAlGuardar.showAndWait();
                }
            } else if (resultado.get() == btnSalirSinGuardar) {
                // Salir sin guardar
                ControladorGeneral.regresar(event, "Principal", getClass());
            }
            // Si se selecciona Cancelar, no hacer nada (permanece en la ventana actual)
        } else {
            // Si no hay cambios, regresar normalmente
            ControladorGeneral.regresar(event, "Principal", getClass());
        }
    }

// Método para verificar si hay cambios sin guardar
    private boolean hayCambiosSinGuardar() {
        // Obtener los estados de acreditación originales
        HashMap<String, Boolean> estadosOriginales = cargarEstadosAcreditacion();

        // Verificar cada evento en la lista actual
        for (Evento evento : tableView.getItems()) {
            // Crear la clave única como en el método cargarEstadosAcreditacion
            String key = evento.getRfc() + "|" + evento.getNombreEvento();

            // Verificar si el estado de acreditación ha cambiado
            Boolean estadoOriginal = estadosOriginales.get(key);
            if (estadoOriginal == null || estadoOriginal != evento.isAcreditado()) {
                return true; // Hay cambios sin guardar
            }
        }

        return false; // No hay cambios
    }

    public static class Evento {

        private String apellidoPaterno, apellidoMaterno, nombres, rfc, sexo, departamento, puesto, nombreEvento, capacitacion, nombreFacilitador, periodo;
        private BooleanProperty acreditacion;

        public Evento(String apellidoPaterno, String apellidoMaterno, String nombres,
                String rfc, String sexo, String departamento, String puesto, String nombreEvento, String capacitacion, String nombreFacilitador,
                String periodo, Boolean acreditacion) {
            this.apellidoPaterno = apellidoPaterno;
            this.apellidoMaterno = apellidoMaterno;
            this.nombres = nombres;
            this.rfc = rfc;
            this.sexo = sexo;
            this.departamento = departamento;
            this.puesto = puesto;
            this.nombreEvento = nombreEvento;
            this.capacitacion = capacitacion;
            this.nombreFacilitador = nombreFacilitador;
            this.periodo = periodo;
            this.acreditacion = new SimpleBooleanProperty(acreditacion);
        }

        public BooleanProperty acreditadoProperty() {
            return acreditacion;
        }

        public void setAcreditado(boolean acreditado) {
            this.acreditacion.set(acreditado);
        }

        public boolean isAcreditado() {
            return acreditacion.get();
        }

        public String getApellidoPaterno() {
            return apellidoPaterno;
        }

        public String getApellidoMaterno() {
            return apellidoMaterno;
        }

        public String getNombres() {
            return nombres;
        }

        public String getRfc() {
            return rfc;
        }

        public String getSexo() {
            return sexo;
        }

        public String getDepartamento() {
            return departamento;
        }

        public String getPuesto() {
            return puesto;
        }

        public String getNombreEvento() {
            return nombreEvento;
        }

        public String getCapacitacion() {
            return capacitacion;
        }

        public void setCapacitacion(String capacitacion) {
            this.capacitacion = capacitacion;
        }

        public String getNombreFacilitador() {
            return nombreFacilitador;
        }

        public String getPeriodo() {
            return periodo;
        }

        public BooleanProperty getAcreditacion() {
            return acreditacion;
        }

        public void setAcreditacion(Boolean acreditacion) {
            this.acreditacion = new SimpleBooleanProperty(acreditacion);
        }

        public Evento() {
            this.acreditacion = new SimpleBooleanProperty(false); // Inicialmente no marcado
        }

    }

    // Modificar el método leerEventosDesdeExcel para usar los estados de acreditación
    public List<Evento> leerEventosDesdeExcel(String rutaArchivoEventos, String rutaArchivoClasificaciones) throws IOException {
        // Cargar estados de acreditación
        HashMap<String, Boolean> estadosAcreditacion = cargarEstadosAcreditacion();

        // Cargar clasificaciones desde el segundo archivo Excel
        HashMap<String, String> clasificaciones = cargarClasificacionesDesdeExcel(rutaArchivoClasificaciones);

        List<Evento> eventos = new ArrayList<>();

        try (FileInputStream archivo = new FileInputStream(rutaArchivoEventos); Workbook workbook = new XSSFWorkbook(archivo)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue; // Saltar la fila de encabezados
                }
                if (row.getCell(4) == null || row.getCell(4).getStringCellValue().trim().isEmpty()) {
                    break;
                }

                String apellidoPaterno = getCellValueSafe(row.getCell(4));
                String apellidoMaterno = getCellValueSafe(row.getCell(5));
                String nombres = getCellValueSafe(row.getCell(6));
                String rfc = getCellValueSafe(row.getCell(7));
                String sexo = getCellValueSafe(row.getCell(8));
                String departamento = getCellValueSafe(row.getCell(9));
                String puesto = getCellValueSafe(row.getCell(10));
                String nombreEvento = getCellValueSafe(row.getCell(11));
                String nombreFacilitador = getCellValueSafe(row.getCell(12));
                String periodo = getCellValueSafe(row.getCell(13));

                // Buscar el estado de acreditación en el mapa
                String key = rfc + "|" + nombreEvento;
                Boolean acreditado = estadosAcreditacion.getOrDefault(key, false); // Por defecto false (tacha)

                String capacitacion = clasificaciones.getOrDefault(
                        nombreEvento.split("\\.")[1].trim(),
                        "Sin clasificación"
                );

                eventos.add(new Evento(
                        apellidoPaterno, apellidoMaterno, nombres, rfc, sexo,
                        departamento, puesto, nombreEvento, capacitacion,
                        nombreFacilitador, periodo, acreditado
                ));
            }
        }

        return eventos;
    }

    // Método para cargar las clasificaciones desde el archivo Excel
    private static HashMap<String, String> cargarClasificacionesDesdeExcel(String rutaArchivoClasificaciones) throws IOException {
        HashMap<String, String> clasificaciones = new HashMap<>();

        try (FileInputStream archivo = new FileInputStream(rutaArchivoClasificaciones); Workbook workbook = new XSSFWorkbook(archivo)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue; // Saltar la fila de encabezados
                }

                String nombreEvento = getCellValue(row.getCell(1)); // Supongamos que columna 0 tiene el nombre del evento
                String tipoCapacitacion = getCellValue(row.getCell(2)); // Supongamos que columna 1 tiene el tipo

                if (!nombreEvento.isEmpty() && !tipoCapacitacion.isEmpty()) {
                    clasificaciones.put(nombreEvento, tipoCapacitacion);
                }
            }
        }
        return clasificaciones;
    }

    // Método para obtener el valor de una celda como String
    private static String getCellValue(Cell cell) {
        return cell != null ? cell.toString().trim() : "";
    }

    // Método para cargar los estados de acreditación desde el archivo condensado
    private HashMap<String, Boolean> cargarEstadosAcreditacion() {
        HashMap<String, Boolean> estadosAcreditacion = new HashMap<>();

        try {
            Calendar calendario = Calendar.getInstance();
            int año = calendario.get(Calendar.YEAR);
            int periodo = (calendario.get(Calendar.MONTH) + 1) < 7 ? 1 : 2;

            String rutaBase = ControladorGeneral.obtenerRutaDeEjecusion()
                    + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos\\"
                    + año + "\\" + periodo + "-" + año + "\\";

            File carpeta = new File(rutaBase);
            if (!carpeta.exists() || !carpeta.isDirectory()) {
                System.out.println("No existe el directorio de condensados");
                return estadosAcreditacion; // Retorna mapa vacío, resultará en todas las tachas
            }

            // Obtener la última versión del condensado
            int numeroVersion = obtenerUltimaSemana(rutaBase, "condensado\\_\\(version_\\d+\\)\\.xlsx", "version");
            String rutaArchivo = rutaBase + "condensado_(version_" + numeroVersion + ").xlsx";

            File archivoCondensado = new File(rutaArchivo);
            if (!archivoCondensado.exists()) {
                System.out.println("No existe archivo condensado");
                return estadosAcreditacion; // Retorna mapa vacío, resultará en todas las tachas
            }

            // Leer el archivo Excel
            try (FileInputStream fis = new FileInputStream(archivoCondensado); Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(0);

                // Encontrar índices de las columnas necesarias
                int[] indices = encontrarIndicesColumnas(sheet.getRow(0));
                if (indices == null) {
                    return estadosAcreditacion; // Si no se encuentran las columnas necesarias
                }

                // Iterar sobre todas las filas
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue; // Saltar encabezados
                    }
                    // Extraer la información necesaria
                    String rfc = getCellValueSafe(row.getCell(indices[3])); // RFC
                    String nombreEvento = getCellValueSafe(row.getCell(indices[7])); // Nombre Evento
                    String acreditacionStr = getCellValueSafe(row.getCell(indices[11])); // Acreditación

                    if (rfc.isEmpty() || nombreEvento.isEmpty()) {
                        continue;
                    }

                    // Crear una clave única combinando RFC y nombre del evento
                    String key = rfc + "|" + nombreEvento;

                    // Determinar el estado de acreditación
                    boolean acreditado = acreditacionStr.toLowerCase().equals("si");

                    // Guardar en el mapa
                    estadosAcreditacion.put(key, acreditado);
                }

            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("Error al leer el archivo condensado: " + e.getMessage());
            }

        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("Error general al cargar estados de acreditación: " + e.getMessage());
        }

        return estadosAcreditacion;
    }

// Método auxiliar para encontrar los índices de las columnas necesarias
    private int[] encontrarIndicesColumnas(Row headerRow) {
        if (headerRow == null) {
            return null;
        }

        int[] indices = new int[12]; // [ApPaterno, ApMaterno, Nombres, RFC, Sexo, Depto, Puesto, NombreEvento, Capacitacion, Facilitador, Periodo, Acreditacion]
        String[] columnasRequeridas = {
            "ApellidoPaterno", "ApellidoMaterno", "Nombres", "RFC", "Sexo",
            "Departamento", "Puesto", "NombreEvento", "Capacitacion",
            "NombreFacilitador", "Periodo", "Acreditación"
        };

        for (Cell cell : headerRow) {
            String headerValue = cell.getStringCellValue().trim();
            for (int i = 0; i < columnasRequeridas.length; i++) {
                if (headerValue.equalsIgnoreCase(columnasRequeridas[i])) {
                    indices[i] = cell.getColumnIndex();
                    break;
                }
            }
        }

        return indices;
    }

// Método auxiliar para obtener el valor de una celda de forma segura
    private String getCellValueSafe(Cell cell) {
        if (cell == null) {
            return "";
        }
        try {
            return cell.getStringCellValue().trim();
        } catch (Exception e) {
            try {
                return String.valueOf(cell.getNumericCellValue()).trim();
            } catch (Exception ex) {
                return "";
            }
        }
    }

}
