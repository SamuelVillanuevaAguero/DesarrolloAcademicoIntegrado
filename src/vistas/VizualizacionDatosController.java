package vistas;

import java.io.File;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Optional;
import java.util.ResourceBundle;
import java.util.logging.Level;
import java.util.logging.Logger;
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
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.CheckBoxTableCell;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utilerias.general.ControladorGeneral;
import utilerias.visualizacionDatos.Evento;
import utilerias.visualizacionDatos.ExcelReader;

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

        configurarTabla();
        cargarDatos();
    }

    private void configurarTabla() {
        TableColumn<Evento, String> colHoraInicio = new TableColumn<>("Fecha de Inicio");
        colHoraInicio.setCellValueFactory(new PropertyValueFactory<>("horaInicio"));

        TableColumn<Evento, String> colHoraFinal = new TableColumn<>("Fecha de Finalización");
        colHoraFinal.setCellValueFactory(new PropertyValueFactory<>("horaFinal"));

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

        TableColumn<Evento, String> colPosgrado = new TableColumn<>("Posgrado");
        colPosgrado.setCellValueFactory(new PropertyValueFactory<>("nivelPosgrado"));

        TableColumn<Evento, String> colPuesto = new TableColumn<>("Puesto");
        colPuesto.setCellValueFactory(new PropertyValueFactory<>("puesto"));

        TableColumn<Evento, String> colNombreEvento = new TableColumn<>("Nombre del Evento");
        colNombreEvento.setCellValueFactory(new PropertyValueFactory<>("nombreEvento"));

        TableColumn<Evento, String> colNombreFacilitador = new TableColumn<>("Nombre del Facilitador");
        colNombreFacilitador.setCellValueFactory(new PropertyValueFactory<>("nombreFacilitador"));

        TableColumn<Evento, String> colPeriodo = new TableColumn<>("Periodo");
        colPeriodo.setCellValueFactory(new PropertyValueFactory<>("periodo"));

        // Columna de botones
        TableColumn<Evento, Void> colAcreditacion = new TableColumn<>("Acreditación");
        colAcreditacion.setCellFactory(getButtonCellFactory());

        tableView.getColumns().addAll(colHoraInicio, colHoraFinal, colApellidoPaterno,
                colApellidoMaterno, colNombres, colRFC, colSexo, colDepartamento, colPosgrado,
                colPuesto, colNombreEvento, colNombreFacilitador, colPeriodo, colAcreditacion);
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
            Calendar calendario = Calendar.getInstance();
            int year = calendario.get(Calendar.YEAR);
            int mesActual = calendario.get(Calendar.MONTH) + 1;

            // Definir la estructura de directorios
            String carpetaPeriodo = (mesActual >= 1 && mesActual <= 7) ? "1-" + year : "2-" + year;
            String rutaBase = ControladorGeneral.obtenerRutaDeEjecusion()
                    + "\\Gestion_de_cursos\\Archivos_importados\\" + year + "\\" + carpetaPeriodo
                    + "\\listado_de_pre_regitro_a_cursos_de_capacitacion\\";

            // Buscar el archivo Excel de la última semana
            Optional<Path> archivoUltimaSemana = obtenerArchivoUltimaSemana(rutaBase);

            if (archivoUltimaSemana.isPresent()) {
                // Leer los datos desde el archivo Excel encontrado
                eventoList.addAll(ExcelReader.leerEventosDesdeExcel(archivoUltimaSemana.get().toString()));
                tableView.setItems(eventoList);
            } else {
                mostrarAlerta("Aviso", "No se encontraron archivos para la última semana.", AlertType.WARNING);
            }

        } catch (IOException e) {
            mostrarAlerta("Error", "No se pudieron cargar los datos del archivo Excel.", AlertType.ERROR);
        }
    }

    /**
     * Obtiene el archivo de Excel correspondiente a la última semana (por fecha
     * de modificación más reciente).
     */
    private Optional<Path> obtenerArchivoUltimaSemana(String rutaDirectorio) {
        try {
            Path directorio = Paths.get(rutaDirectorio);

            // Validar si el directorio existe
            if (!Files.exists(directorio) || !Files.isDirectory(directorio)) {
                return Optional.empty();
            }

            // Filtrar los archivos con el patrón `listado_(Semana_X).xlsx`
            List<Path> archivosSemana = Files.list(directorio)
                    .filter(archivo -> archivo.getFileName().toString().matches("listado_\\(Semana_\\d+\\)\\.xlsx"))
                    .sorted((a, b) -> {
                        // Ordenar por fecha de modificación (descendente)
                        try {
                            return Files.getLastModifiedTime(b).compareTo(Files.getLastModifiedTime(a));
                        } catch (IOException e) {
                            return 0; // Si hay un error, considerar iguales
                        }
                    })
                    .collect(Collectors.toList());

            // Retornar el primero de la lista (el más reciente)
            return archivosSemana.isEmpty() ? Optional.empty() : Optional.of(archivosSemana.get(0));

        } catch (IOException e) {
            e.printStackTrace();
            return Optional.empty();
        }
    }

    @FXML
    private void Exportar(ActionEvent event) {
        // Crear un FileChooser para seleccionar dónde guardar el archivo
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Guardar archivo");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Archivos CSV", "ListaAsistencia.csv"));

        // Abrir el diálogo de guardar
        var file = fileChooser.showSaveDialog(null);
        if (file != null) {
            try (FileWriter writer = new FileWriter(file)) {
                // Escribir encabezados en el archivo CSV
                writer.write("Nombre del Participante, Curso\n");

                // Recorrer las filas de la tabla
                List<Evento> participantes = tableView.getItems();
                for (Evento participante : participantes) {
                    // Filtrar por curso y acreditación
                    if (participante.getNombreEvento().equals(campoBusqueda) && participante.isAcreditado()) {
                        writer.write(participante.getNombres() + ", " + participante.getNombreEvento() + "\n");
                    }
                }

                System.out.println("Archivo exportado con éxito");
            } catch (IOException e) {
                System.out.println("Error al exportar: " + e.getMessage());
            }
        }
    }

    private void mostrarAlerta(String titulo, String mensaje, AlertType tipo) {
        Alert alert = new Alert(tipo);
        alert.setTitle(titulo);
        alert.setContentText(mensaje);
        alert.showAndWait();
    }

    @FXML
    private void Buscar(ActionEvent event) {
        String textoBusqueda = campoBusqueda.getText().trim().toLowerCase();

        // Si el campo de búsqueda está vacío, mostramos una advertencia.
        if (textoBusqueda.isEmpty()) {
            Alert alert = new Alert(Alert.AlertType.WARNING);
            alert.setTitle("Campo de Búsqueda Vacío");
            alert.setHeaderText(null);
            alert.setContentText("Por favor, introduce un valor para buscar.");
            alert.showAndWait();
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
                    || info.getHoraInicio().toLowerCase().contains(textoBusqueda)
                    || info.getHoraFinal().toLowerCase().contains(textoBusqueda)
                    || info.getNombreEvento().toLowerCase().contains(textoBusqueda)
                    || info.getDepartamento().toLowerCase().contains(textoBusqueda)
                    || info.getNivelPosgrado().toLowerCase().contains(textoBusqueda)
                    || info.getNombreFacilitador().toLowerCase().contains(textoBusqueda)
                    || info.getPeriodo().toLowerCase().contains(textoBusqueda)
                    || info.getPuesto().toLowerCase().contains(textoBusqueda);
        });

        // Si no se encuentra ningún resultado, mostramos un mensaje de información y restauramos la tabla completa.
        if (datosFiltrados.isEmpty()) {
            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.setTitle("Resultado de Búsqueda");
            alert.setHeaderText(null);
            alert.setContentText("No se encontró ningún resultado para \"" + textoBusqueda + "\".");
            alert.showAndWait();

            // Restauramos la tabla con todos los datos
            tableView.setItems(eventoList);
        } else {
            // Si hay resultados, mostramos solo los datos filtrados.
            tableView.setItems(datosFiltrados);
        }

        // Limpiar el campo de búsqueda después de realizar la búsqueda
        campoBusqueda.clear();
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
                "HoraInicio", "HoraFinal", "ApellidoPaterno", "ApellidoMaterno", "Nombres",
                "RFC", "Sexo", "Departamento", "Posgrado", "Puesto", "NombreEvento",
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
                row.createCell(0).setCellValue(evento.getHoraInicio());
                row.createCell(1).setCellValue(evento.getHoraFinal());
                row.createCell(2).setCellValue(evento.getApellidoPaterno());
                row.createCell(3).setCellValue(evento.getApellidoMaterno());
                row.createCell(4).setCellValue(evento.getNombres());
                row.createCell(5).setCellValue(evento.getRfc());
                row.createCell(6).setCellValue(evento.getSexo());
                row.createCell(7).setCellValue(evento.getDepartamento());
                row.createCell(8).setCellValue(evento.getNivelPosgrado());
                row.createCell(9).setCellValue(evento.getPuesto());
                row.createCell(10).setCellValue(evento.getNombreEvento());
                row.createCell(11).setCellValue(evento.getNombreFacilitador());
                row.createCell(12).setCellValue(evento.getPeriodo());
                row.createCell(13).setCellValue(evento.isAcreditado() ? "si" : "no");
            }

            // Autoajustar el ancho de las columnas
            for (int i = 0; i < columnas.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Obtener el año y la carpeta del período
            Calendar calendario = Calendar.getInstance();
            int year = calendario.get(Calendar.YEAR);
            int mesActual = calendario.get(Calendar.MONTH) + 1;

            String carpetaPeriodo = (mesActual >= 1 && mesActual <= 7) ? "1-" + year : "2-" + year;

            // Definir la carpeta donde se guardarán los archivos
            String carpetaDestino = ControladorGeneral.obtenerRutaDeEjecusion()
                    + "\\Gestion_de_Cursos\\Sistema\\condensados_vista_de_visualizacion_de_datos\\" + year + "\\" + carpetaPeriodo;

            // Crear la carpeta si no existe
            File carpeta = new File(carpetaDestino);
            if (!carpeta.exists()) {
                carpeta.mkdirs();
            }

            // Determinar el número de versión
            int numeroDeVersion = determinarNumeroDeVersion(carpetaDestino);

            // Definir la ruta del archivo con el número de versión
            String filePath = carpetaDestino + "\\condensado_(version_" + numeroDeVersion + ").xlsx";

            // Guardar el archivo en la ruta definida
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

    /**
     * Determina el número de versión para el siguiente archivo. Busca los
     * archivos existentes en la carpeta y calcula el siguiente número de
     * versión.
     */
    private int determinarNumeroDeVersion(String carpetaDestino) {
        File carpeta = new File(carpetaDestino);
        File[] archivos = carpeta.listFiles((dir, name) -> name.matches("condensado_\\(version_\\d+\\)\\.xlsx"));

        if (archivos == null || archivos.length == 0) {
            return 1; // Si no hay archivos, la primera versión es 1
        }

        // Buscar el número de la versión más alta
        int maxVersion = 0;
        for (File archivo : archivos) {
            String nombre = archivo.getName();
            String numero = nombre.replaceAll("[^0-9]", ""); // Extraer el número de versión
            try {
                maxVersion = Math.max(maxVersion, Integer.parseInt(numero));
            } catch (NumberFormatException e) {
                // Ignorar nombres que no sigan el patrón esperado
            }
        }

        return maxVersion + 1; // Retornar la siguiente versión disponible
    }

    //Métodos de los botones de la barra superior :)
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
