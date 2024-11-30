/*
     * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
     * Click nbfs://nbhost/SystemFileSystem/Templates/javafx/FXMLController.java to edit this template
 */
package vistas;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.net.URL;
import java.time.Year;
import java.util.Calendar;
import java.util.HashMap;
import java.util.ResourceBundle;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.Stage;
import utilerias.general.ControladorGeneral;
import java.util.Optional;

public class ModificacionDatosController implements Initializable {

    @FXML
    private Button botonCerrar;
    @FXML
    private Button botonMinimizar;
    @FXML
    private Button botonRegresar;
    @FXML
    private Button botonGuardar;
    @FXML
    private Button botonLimpiar;
    @FXML
    private TextField directorField;
    @FXML
    private TextField coordinadorField;
    @FXML
    private TextField jefeDeptoField;
    @FXML
    private Spinner<Integer> totalDocentes;
    @FXML
    private Spinner<Integer> año;
    @FXML
    private ComboBox<String> periodoEscolar;
    @FXML
    private ComboBox<String> departamento;

    private HashMap<String, Integer> docentesPorDepartamento = new HashMap<>();

    public void regresarVentana(MouseEvent event) throws IOException {
        // Verificar si hay datos en los campos
        boolean hayDatos = !directorField.getText().trim().isEmpty()
                || !coordinadorField.getText().trim().isEmpty()
                || !jefeDeptoField.getText().trim().isEmpty()
                || (totalDocentes.getValue() != null && totalDocentes.getValue() != 100)
                || periodoEscolar.getValue() != null;
        if (hayDatos) {
            Alert confirmacion = new Alert(Alert.AlertType.CONFIRMATION);
            confirmacion.setTitle("Confirmación");
            confirmacion.setHeaderText("¿Deseas guardar los datos?");
            confirmacion.setContentText("Tienes datos ingresados en los campos. ¿Deseas guardarlos antes de regresar?");

            ButtonType botonGuardar = new ButtonType("Guardar");
            ButtonType botonNoGuardar = new ButtonType("No Guardar");
            ButtonType botonCancelar = new ButtonType("Cancelar", ButtonBar.ButtonData.CANCEL_CLOSE);

            confirmacion.getButtonTypes().setAll(botonGuardar, botonNoGuardar, botonCancelar);
            Optional<ButtonType> resultado = confirmacion.showAndWait();

            if (resultado.isPresent()) {
                if (resultado.get() == botonGuardar) {
                    guardarDatosEnExcel(event);
                    ControladorGeneral.regresar(event, "Principal", getClass());
                } else if (resultado.get() == botonNoGuardar) {
                    ControladorGeneral.regresar(event, "Principal", getClass());
                }
            }
        } else {
            ControladorGeneral.regresar(event, "Principal", getClass());
        }
    }

    // Método para mostrar alertas en la aplicación
    private void mostrarAlerta(String titulo, String mensaje, AlertType tipo) {
        Alert alerta = new Alert(tipo);
        alerta.setTitle(titulo);
        alerta.setHeaderText(null);
        alerta.setContentText(mensaje);
        alerta.showAndWait();
    }

    //Métodos de los botones de la barra superior :)
    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    // Nuevo método para restablecer campos a valores iniciales
    public void restablecerCamposValoresIniciales() {
        directorField.clear();
        coordinadorField.clear();
        jefeDeptoField.clear();
        totalDocentes.getValueFactory().setValue(100); // Valor inicial del Spinner
        año.getValueFactory().setValue(Calendar.getInstance().get(Calendar.YEAR)); // Valor inicial del año
        periodoEscolar.getSelectionModel().clearSelection(); // Limpiar selección del ComboBox
    }

    // Método auxiliar para sumar docentes en un rango de filas
    private int sumarDocentes(Sheet sheet, int filaInicio, int filaFin) {
        int suma = 0;
        for (int i = filaInicio; i <= filaFin; i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getCell(1) != null) {
                suma += row.getCell(1).getNumericCellValue();
            }
        }
        return suma;
    }

    public void guardarDatosEnExcel(MouseEvent event) {
        if (!validarCampos()) {
            return;
        }

        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx";
        File archivo = new File(rutaArchivo);
        Workbook workbook;

        try {
            // Verificar si el archivo existe
            if (archivo.exists()) {
                // Si existe, cargar el archivo existente
                workbook = new XSSFWorkbook(new FileInputStream(archivo));
            } else {
                // Si no existe, crear uno nuevo
                workbook = new XSSFWorkbook();
            }

            // Obtener o crear la hoja para el año seleccionado
            String nombreHoja = String.valueOf(año.getValue());
            Sheet sheet = workbook.getSheet(nombreHoja);
            if (sheet == null) {
                sheet = workbook.createSheet(nombreHoja);
            }

            // Configuración de cabeceras si es una hoja nueva
            if (sheet.getPhysicalNumberOfRows() == 0) {
                sheet.createRow(0).createCell(0).setCellValue("Director:");
                sheet.createRow(1).createCell(0).setCellValue("Jefe de Departamento:");
                sheet.createRow(2).createCell(0).setCellValue("Coordinador:");
                sheet.createRow(3).createCell(0).setCellValue("Periodo Enero-Julio:");
                sheet.createRow(12).createCell(0).setCellValue("Periodo Agosto-Diciembre:");
            }

            // Actualizar datos básicos solo si los campos no están vacíos
            if (!directorField.getText().trim().isEmpty()) {
                sheet.getRow(0).createCell(1).setCellValue(directorField.getText());
            }
            if (!jefeDeptoField.getText().trim().isEmpty()) {
                sheet.getRow(1).createCell(1).setCellValue(jefeDeptoField.getText());
            }
            if (!coordinadorField.getText().trim().isEmpty()) {
                sheet.getRow(2).createCell(1).setCellValue(coordinadorField.getText());
            }

            // Obtener el periodo seleccionado
            String periodoSeleccionado = periodoEscolar.getValue();
            int filaInicio = periodoSeleccionado.equals("Enero-Julio") ? 4 : 13; // Fila base según el periodo

            // Guardar docentes por departamento en el periodo correspondiente
            int fila = filaInicio;
            for (String depto : departamento.getItems()) {
                Row row = sheet.getRow(fila);
                if (row == null) {
                    row = sheet.createRow(fila);
                }
                row.createCell(0).setCellValue(depto); // Nombre del departamento
                row.createCell(1).setCellValue(docentesPorDepartamento.getOrDefault(depto, 100)); // Número de docentes
                fila++;
            }
            // Calcular y escribir la suma total de docentes en los periodos
            int sumaEneroJulio = sumarDocentes(sheet, 4, 12); // Filas del periodo Enero-Julio
            int sumaAgostoDiciembre = sumarDocentes(sheet, 13, 21); // Filas del periodo Agosto-Diciembre

            sheet.getRow(3).createCell(1).setCellValue(sumaEneroJulio); // Total en Enero-Julio
            sheet.getRow(12).createCell(1).setCellValue(sumaAgostoDiciembre); // Total en Agosto-Diciembre

            // Ajustar el ancho de las columnas automáticamente
            for (int i = 0; i <= 1; i++) {
                sheet.autoSizeColumn(i);
            }

            // Guardar archivo
            try (FileOutputStream fileOut = new FileOutputStream(rutaArchivo)) {
                workbook.write(fileOut);
                mostrarAlerta("Éxito", "Datos guardados exitosamente", AlertType.INFORMATION);

                // Limpiar campos después de guardar exitosamente
                restablecerCamposValoresIniciales();
            }
        } catch (IOException e) {
            mostrarAlerta("Error", "No se pudo guardar el archivo: " + e.getMessage(), AlertType.ERROR);
            e.printStackTrace();
        }
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {

        departamento.getItems().addAll(
                "CIENCIAS BASICAS",
                "CIENCIAS ECONOMICO ADMINISTRATIVO",
                "CIENCIAS DE LA TIERRA",
                "INGENIERIA INDUSTRIAL",
                "METAL MECANICA",
                "QUIMICA Y BIOQUIMICA",
                "SISTEMAS COMPUTACIONALES",
                "POSGRADO"
        );

        // Inicializar valores por defecto en el HashMap
        for (String depto : departamento.getItems()) {
            docentesPorDepartamento.put(depto, 10); // Por defecto, 10 docentes
        }

        // Configuración del Spinner de docentes
        SpinnerValueFactory<Integer> docentesFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(10, 100, 10);
        totalDocentes.setValueFactory(docentesFactory);
        totalDocentes.setEditable(true);

        // Configuración del Spinner de año
        int añoActual = Calendar.getInstance().get(Calendar.YEAR); // Año actual fijo para el sistema
        SpinnerValueFactory<Integer> añoFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(2000, 2100, añoActual);
        año.setValueFactory(añoFactory);
        año.setEditable(true);

        // Actualizar el Spinner de docentes al seleccionar un departamento
        departamento.getSelectionModel().selectedItemProperty().addListener((observable, oldValue, newValue) -> {
            if (newValue != null) {
                totalDocentes.getValueFactory().setValue(docentesPorDepartamento.get(newValue));
            }
        });

        // Guardar cambios en el HashMap cuando se edita el Spinner
        totalDocentes.valueProperty().addListener((observable, oldValue, newValue) -> {
            String deptoSeleccionado = departamento.getValue();
            if (deptoSeleccionado != null) {
                docentesPorDepartamento.put(deptoSeleccionado, newValue);
            }
        });

        // Validación corregida para el Spinner de año
        año.getEditor().focusedProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue) { // Cuando pierde el foco
                try {
                    String text = año.getEditor().getText();
                    int value;

                    if (text.isEmpty()) {
                        año.getValueFactory().setValue(añoActual);
                    } else {
                        value = Integer.parseInt(text);
                        if (value < 2000 || value > 2100) {
                            mostrarAlerta("Validación", "El año debe estar entre 2000 y 2100", AlertType.WARNING);
                            año.getValueFactory().setValue(añoActual);

                        } else {
                            año.getValueFactory().setValue(value);
                        }
                    }
                } catch (NumberFormatException e) {
                    año.getValueFactory().setValue(añoActual);
                }
            }
        });

        // Solo permitir números en los Spinners
        totalDocentes.getEditor().textProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue.matches("\\d*")) {
                totalDocentes.getEditor().setText(oldValue);
            }
        });

        // Solo permitir números en el Spinner de año
        año.getEditor().textProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue.matches("\\d*")) {
                año.getEditor().setText(oldValue);
            } else if (!newValue.isEmpty()) {
                try {
                    int value = Integer.parseInt(newValue);
                    if (value > 2100) {
                        año.getEditor().setText(oldValue);
                    }
                } catch (NumberFormatException e) {
                    año.getEditor().setText(oldValue);
                }
            }
        });

        periodoEscolar.getItems().addAll("Enero-Julio", "Agosto-Diciembre");

        // Configurar eventos de botones
        botonGuardar.setOnMouseClicked(this::guardarDatosEnExcel);
        botonCerrar.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión??", getClass()
                );
            } catch (IOException ex) {
                Logger.getLogger(ModificacionDatosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        botonMinimizar.setOnMouseClicked(this::minimizarVentana);
        botonRegresar.setOnMouseClicked(event -> {
            try {
                regresarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(ModificacionDatosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        botonLimpiar.setOnMouseClicked(event -> restablecerCamposValoresIniciales());
    }

    private boolean validarCampos() {
        // Verificar si todos los campos están vacíos
        boolean todosVacios = directorField.getText().trim().isEmpty()
                && coordinadorField.getText().trim().isEmpty()
                && jefeDeptoField.getText().trim().isEmpty()
                && periodoEscolar.getValue() == null;

        if (todosVacios) {
            mostrarAlerta("Validación", "Todos los campos son obligatorios. Por favor, complete el formulario.", AlertType.WARNING);
            return false;
        }

        // Validaciones individuales
        if (directorField.getText().trim().isEmpty()) {
            mostrarAlerta("Validación", "Por favor, ingrese el nombre del Director.", AlertType.WARNING);
            directorField.requestFocus();
            return false;
        }

        if (coordinadorField.getText().trim().isEmpty()) {
            mostrarAlerta("Validación", "Por favor, ingrese el nombre del Coordinador.", AlertType.WARNING);
            coordinadorField.requestFocus();
            return false;
        }

        if (jefeDeptoField.getText().trim().isEmpty()) {
            mostrarAlerta("Validación", "Por favor, ingrese el nombre del Jefe de Departamento.", AlertType.WARNING);
            jefeDeptoField.requestFocus();
            return false;
        }

        if (periodoEscolar.getValue() == null) {
            mostrarAlerta("Validación", "Por favor, seleccione un periodo escolar.", AlertType.WARNING);
            periodoEscolar.requestFocus();
            return false;
        }

        // Validar formato de nombres solo si no están vacíos
        if (!directorField.getText().matches("[a-zA-ZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Director' solo debe contener letras.", AlertType.WARNING);
            directorField.requestFocus();
            return false;
        }
        if (!coordinadorField.getText().matches("[a-zA-ZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Coordinador' solo debe contener letras.", AlertType.WARNING);
            coordinadorField.requestFocus();
            return false;
        }
        if (!jefeDeptoField.getText().matches("[a-zA-ZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Jefe de Departamento' solo debe contener letras.", AlertType.WARNING);
            jefeDeptoField.requestFocus();
            return false;
        }

        return true;
    }
}
