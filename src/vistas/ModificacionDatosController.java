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
import javafx.application.Platform;

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
                || (totalDocentes.getValue() != null && totalDocentes.getValue() != 0)
                || periodoEscolar.getValue() != null;
        if (hayDatos) {
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
                guardarDatosEnExcel(event);
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

    public void restablecerCamposValoresIniciales() {
        // Modificar para permitir limpiar campos siempre
        directorField.clear();
        coordinadorField.clear();
        jefeDeptoField.clear();

        for (String depto : departamento.getItems()) {
            docentesPorDepartamento.put(depto, 0);
        }

        totalDocentes.getValueFactory().setValue(0);
        año.getValueFactory().setValue(Calendar.getInstance().get(Calendar.YEAR));
        periodoEscolar.getSelectionModel().clearSelection();
        departamento.getSelectionModel().clearSelection();

        // Volver a deshabilitar campos
        deshabilitarCampos();
    }

    private boolean noHayRegistrosPrevios() {
        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx";
        File archivo = new File(rutaArchivo);

        if (!archivo.exists()) {
            return true;
        }

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(archivo))) {
            Sheet sheet = workbook.getSheet(String.valueOf(año.getValue()));
            return sheet == null || sheet.getPhysicalNumberOfRows() <= 4; // Considerar solo las filas con datos
        } catch (IOException e) {
            return true;
        }
    }

    public void guardarDatosEnExcel(MouseEvent event) {
        if (!validarCampos()) {
            return;
        }

        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx";
        File archivo = new File(rutaArchivo);

        try {
            Workbook workbook;
            Sheet sheet;

            if (archivo.exists()) {
                workbook = new XSSFWorkbook(new FileInputStream(archivo));
                sheet = workbook.getSheet(String.valueOf(año.getValue()));

                if (sheet == null) {
                    sheet = workbook.createSheet(String.valueOf(año.getValue()));
                    configurarCabecerasHoja(sheet);
                }
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet(String.valueOf(año.getValue()));
                configurarCabecerasHoja(sheet);
            }

            // Actualizar datos en la hoja
            actualizarDatosEnHoja(sheet);

            // Guardar archivo
            try (FileOutputStream fileOut = new FileOutputStream(rutaArchivo)) {
                workbook.write(fileOut);
                mostrarAlerta("Éxito", "Datos guardados exitosamente", AlertType.INFORMATION);

                // Después de guardar, volver a cargar la información del periodo
                //cargarInformacionPeriodo(periodoEscolar.getValue());
                restablecerCamposValoresIniciales();
            }
        } catch (IOException e) {
            mostrarAlerta("Error", "No se pudo guardar el archivo: " + e.getMessage(), AlertType.ERROR);
            e.printStackTrace();
        }
    }

    private void cargarDatosExistentes(Sheet sheet) {
        // Obtener el periodo seleccionado
        String periodoSeleccionado = periodoEscolar.getValue();

        // Configurar rangos de filas según el periodo
        int filaInicial, filaTotal;
        if (periodoSeleccionado.equals("Enero-Julio")) {
            filaInicial = 5;
            filaTotal = 12;
        } else { // Agosto-Diciembre
            filaInicial = 14;
            filaTotal = 21;
        }

        try {
            // Cargar datos de las primeras filas (siempre iguales)
            Row directorRow = sheet.getRow(0);
            if (directorRow != null && directorRow.getCell(1) != null) {
                directorField.setText(directorRow.getCell(1).getStringCellValue());
            }

            Row jefeDeptoRow = sheet.getRow(1);
            if (jefeDeptoRow != null && jefeDeptoRow.getCell(1) != null) {
                jefeDeptoField.setText(jefeDeptoRow.getCell(1).getStringCellValue());
            }

            Row coordinadorRow = sheet.getRow(2);
            if (coordinadorRow != null && coordinadorRow.getCell(1) != null) {
                coordinadorField.setText(coordinadorRow.getCell(1).getStringCellValue());
            }

            // Limpiar mapa de docentes
            docentesPorDepartamento.clear();
            int totalDocentesValor = 0;

            // Cargar datos de departamentos para el periodo seleccionado
            for (int i = 0; i < departamento.getItems().size(); i++) {
                String depto = departamento.getItems().get(i);
                Row row = sheet.getRow(filaInicial + i);

                if (row != null && row.getCell(1) != null) {
                    int docentes = (int) row.getCell(1).getNumericCellValue();
                    docentesPorDepartamento.put(depto, docentes);
                    totalDocentesValor += docentes;
                } else {
                    docentesPorDepartamento.put(depto, 0);
                }
            }

            // Establecer valores
            totalDocentes.getValueFactory().setValue(totalDocentesValor);

        } catch (Exception e) {
            mostrarAlerta("Error", "No se pudieron cargar los datos existentes: " + e.getMessage(), AlertType.ERROR);
            e.printStackTrace();
        }
    }

    private void actualizarDatosEnHoja(Sheet sheet) {
        // Obtener el periodo seleccionado
        String periodoSeleccionado = periodoEscolar.getValue();

        // Actualizar datos generales (filas 0-2) - estas son comunes
        Row directorRow = sheet.getRow(0);
        if (directorRow.getCell(1) == null) {
            directorRow.createCell(1);
        }
        directorRow.getCell(1).setCellValue(directorField.getText());

        Row jefeDeptoRow = sheet.getRow(1);
        if (jefeDeptoRow.getCell(1) == null) {
            jefeDeptoRow.createCell(1);
        }
        jefeDeptoRow.getCell(1).setCellValue(jefeDeptoField.getText());

        Row coordinadorRow = sheet.getRow(2);
        if (coordinadorRow.getCell(1) == null) {
            coordinadorRow.createCell(1);
        }
        coordinadorRow.getCell(1).setCellValue(coordinadorField.getText());

        // Configurar rangos según periodo
        int filaInicial, filaTotal, filaSuma;
        if (periodoSeleccionado.equals("Enero-Julio")) {
            filaInicial = 4;
            filaTotal = 12;
            filaSuma = 3;
        } else { // Agosto-Diciembre
            filaInicial = 13;
            filaTotal = 21;
            filaSuma = 12;
        }

        // Calcular suma de departamentos para el periodo seleccionado
        int sumaDocentes = 0;
        for (int i = 0; i < 8; i++) {
            String depto = departamento.getItems().get(i);
            int valorDocentes = docentesPorDepartamento.get(depto);
            Row row = sheet.getRow(filaInicial + i);

            if (row.getCell(1) == null) {
                row.createCell(1);
            }
            row.getCell(1).setCellValue(valorDocentes);

            sumaDocentes += valorDocentes;
        }

        // Actualizar suma en la fila correspondiente
        Row sumaRow = sheet.getRow(filaSuma);
        if (sumaRow.getCell(1) == null) {
            sumaRow.createCell(1);
        }
        sumaRow.getCell(1).setCellValue(sumaDocentes);
    }

    private void configurarCabecerasHoja(Sheet sheet) {
        // Establecer cabeceras fijas en las primeras filas
        Row directorRow = sheet.createRow(0);
        directorRow.createCell(0).setCellValue("Director:");
        directorRow.createCell(1).setCellValue(""); // Celda para el nombre

        Row jefeDeptoRow = sheet.createRow(1);
        jefeDeptoRow.createCell(0).setCellValue("Jefe de Departamento:");
        jefeDeptoRow.createCell(1).setCellValue(""); // Celda para el nombre

        Row coordinadorRow = sheet.createRow(2);
        coordinadorRow.createCell(0).setCellValue("Coordinador:");
        coordinadorRow.createCell(1).setCellValue(""); // Celda para el nombre

        // Crear fila para periodo Enero-Julio con suma
        Row periodoEneroJulioRow = sheet.createRow(3);
        periodoEneroJulioRow.createCell(0).setCellValue("Periodo Enero-Julio:");
        periodoEneroJulioRow.createCell(1).setCellValue(0); // Suma inicial en 0

        // Crear filas para departamentos Enero-Julio
        String[] departamentos = {
            "CIENCIAS BASICAS",
            "CIENCIAS ECONOMICO ADMINISTRATIVO",
            "CIENCIAS DE LA TIERRA",
            "INGENIERIA INDUSTRIAL",
            "METAL MECANICA",
            "QUIMICA Y BIOQUIMICA",
            "SISTEMAS COMPUTACIONALES",
            "POSGRADO"
        };

        for (int i = 0; i < departamentos.length; i++) {
            Row row = sheet.createRow(i + 4);
            row.createCell(0).setCellValue(departamentos[i]);
            row.createCell(1).setCellValue(0); // Valor inicial en 0
        }

        // Crear fila para periodo Agosto-Diciembre con suma
        Row periodoAgostoDiciembreRow = sheet.createRow(12);
        periodoAgostoDiciembreRow.createCell(0).setCellValue("Periodo Agosto-Diciembre:");
        periodoAgostoDiciembreRow.createCell(1).setCellValue(0); // Suma inicial en 0

        // Crear filas para departamentos Agosto-Diciembre
        for (int i = 0; i < departamentos.length; i++) {
            Row row = sheet.createRow(i + 13);
            row.createCell(0).setCellValue(departamentos[i]);
            row.createCell(1).setCellValue(0); // Valor inicial en 0
        }
    }

    private boolean validarCampos() {
        // Campos obligatorios siempre
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

        // Validar que cada departamento tenga un valor entre 10 y 80
        boolean departamentosValidos = true;
        StringBuilder mensaje = new StringBuilder("Los siguientes departamentos requieren un valor entre 10 y 80:\n");

        for (String depto : departamento.getItems()) {
            int valorDocentes = docentesPorDepartamento.get(depto);
            if (valorDocentes < 10 || valorDocentes > 80) {
                mensaje.append("- ").append(depto).append("\n");
                departamentosValidos = false;
            }
        }

        if (!departamentosValidos) {
            mostrarAlerta("Validación", mensaje.toString(), AlertType.WARNING);
            return false;
        }

        // Validaciones de formato de nombre
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

    private void verificarArchivoPorAño() {
        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx";
        File archivo = new File(rutaArchivo);

        if (!archivo.exists()) {
            habilitarCampos();
            periodoEscolar.setDisable(false);
            return;
        }

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(archivo))) {
            Sheet sheet = workbook.getSheet(String.valueOf(año.getValue()));

            if (sheet == null) {
                habilitarCampos();
                periodoEscolar.setDisable(false);
            } else {
                periodoEscolar.setDisable(false);
            }
        } catch (IOException e) {
            mostrarAlerta("Error", "No se pudo acceder al archivo: " + e.getMessage(), Alert.AlertType.ERROR);
        }
    }

    private void deshabilitarCampos() {
        directorField.setDisable(true);
        coordinadorField.setDisable(true);
        jefeDeptoField.setDisable(true);
        totalDocentes.setDisable(true);
        departamento.setDisable(true);
        botonGuardar.setDisable(true);
        botonLimpiar.setDisable(true);
    }

    private void habilitarCampos() {
        directorField.setDisable(false);
        coordinadorField.setDisable(false);
        jefeDeptoField.setDisable(false);
        totalDocentes.setDisable(false);
        departamento.setDisable(false);
        botonGuardar.setDisable(false);
        botonLimpiar.setDisable(false);
    }

    private void cargarInformacionPeriodo(String periodoSeleccionado) {
        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx";
        File archivo = new File(rutaArchivo);

        if (!archivo.exists()) {
            habilitarCampos();
            return;
        }

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(archivo))) {
            Sheet sheet = workbook.getSheet(String.valueOf(año.getValue()));

            if (sheet == null) {
                habilitarCampos();
                return;
            }

            // Cargar datos generales
            Row directorRow = sheet.getRow(0);
            Row jefeDeptoRow = sheet.getRow(1);
            Row coordinadorRow = sheet.getRow(2);

            if (directorRow != null && directorRow.getCell(1) != null) {
                directorField.setText(directorRow.getCell(1).getStringCellValue());
            }

            if (jefeDeptoRow != null && jefeDeptoRow.getCell(1) != null) {
                jefeDeptoField.setText(jefeDeptoRow.getCell(1).getStringCellValue());
            }

            if (coordinadorRow != null && coordinadorRow.getCell(1) != null) {
                coordinadorField.setText(coordinadorRow.getCell(1).getStringCellValue());
            }

            // Configurar rangos según periodo
            int filaInicial, filaFinal;
            if (periodoSeleccionado.equals("Enero-Julio")) {
                filaInicial = 4;
                filaFinal = 12;
            } else {
                filaInicial = 13;
                filaFinal = 21;
            }

            // Cargar totales por departamento
            int totalDocentesPeriodo = 0;
            docentesPorDepartamento.clear();

            // Find the first department with a non-zero value
            String primerDepartamentoConValor = null;
            for (int i = 0; i < departamento.getItems().size(); i++) {
                String depto = departamento.getItems().get(i);
                Row row = sheet.getRow(filaInicial + i);

                if (row != null && row.getCell(1) != null) {
                    int docentes = (int) row.getCell(1).getNumericCellValue();
                    docentesPorDepartamento.put(depto, docentes);
                    totalDocentesPeriodo += docentes;

                    // Store the first department with a non-zero value
                    if (primerDepartamentoConValor == null && docentes > 0) {
                        primerDepartamentoConValor = depto;
                    }
                } else {
                    docentesPorDepartamento.put(depto, 0);
                }
            }

            totalDocentes.getValueFactory().setValue(totalDocentesPeriodo);

            // If a department with value is found, select it
            if (primerDepartamentoConValor != null) {
                departamento.setValue(primerDepartamentoConValor);
                totalDocentes.getValueFactory().setValue(docentesPorDepartamento.get(primerDepartamentoConValor));
            }

            // Habilitar campos
            habilitarCampos();

        } catch (IOException e) {
            mostrarAlerta("Error", "No se pudo cargar la información: " + e.getMessage(), Alert.AlertType.ERROR);
        }
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        // Inicializar estados iniciales
        deshabilitarCampos();

        // Establecer foco inicial en año
        Platform.runLater(() -> año.requestFocus());

        // Listener para quitar foco de año
        año.focusedProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue) {
                verificarArchivoPorAño();
            }
        });

        // Listener para selección de periodo
        periodoEscolar.getSelectionModel().selectedItemProperty().addListener((observable, oldValue, newValue) -> {
            if (newValue != null) {
                cargarInformacionPeriodo(newValue);
            }
        });

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
            docentesPorDepartamento.put(depto, 0); // Inicializar en 0
        }

        // Configuración del Spinner de docentes
        SpinnerValueFactory<Integer> docentesFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 80, 0);
        totalDocentes.setValueFactory(docentesFactory);
        totalDocentes.setEditable(true);

        // Configuración del Spinner de año
        int añoActual = Calendar.getInstance().get(Calendar.YEAR);
        SpinnerValueFactory<Integer> añoFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(2000, 2100, añoActual);
        año.setValueFactory(añoFactory);
        año.setEditable(true);

        // Actualizar el Spinner de docentes al seleccionar un departamento
        departamento.getSelectionModel().selectedItemProperty().addListener((observable, oldValue, newValue) -> {
            if (newValue != null) {
                int valorDocentes = docentesPorDepartamento.getOrDefault(newValue, 0);
                totalDocentes.getValueFactory().setValue(valorDocentes);
            }
        });

        // Guardar cambios en el HashMap cuando se edita el Spinner
        totalDocentes.valueProperty().addListener((observable, oldValue, newValue) -> {
            String deptoSeleccionado = departamento.getValue();
            if (deptoSeleccionado != null) {
                docentesPorDepartamento.put(deptoSeleccionado, newValue);
            }
        });

        periodoEscolar.getItems().addAll("Enero-Julio", "Agosto-Diciembre");

        // Configurar eventos de botones
        botonGuardar.setOnMouseClicked(this::guardarDatosEnExcel);
        botonCerrar.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión??", getClass());
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

        // El botón de limpiar no está disponible cuando hay registros previos
        //botonLimpiar.setDisable(!noHayRegistrosPrevios());
        botonLimpiar.setOnMouseClicked(event -> restablecerCamposValoresIniciales());
    }
}
