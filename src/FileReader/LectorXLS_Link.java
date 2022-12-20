package FileReader;

import FileWriter.EscritorTxt;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import static java.lang.String.valueOf;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class LectorXLS_Link {

    private String path;
    private FileInputStream inputStream;
    private Workbook wb;
    private Sheet hoja;
    private EscritorTxt escritor;
    private EscritorTxt escritorControl;
    private DataFormatter formatoDeDatos;
    private int contadorFilas = 2;
    private double montoTotalPrimerVto = 0;
    private double montoTotalSegundoVto = 0;
    private DecimalFormat df;

    public LectorXLS_Link(String path) {

        try {

            this.path = path;
            this.inputStream = new FileInputStream(new File(path));
            this.wb = WorkbookFactory.create(inputStream);
            this.hoja = wb.getSheetAt(0);
            this.formatoDeDatos = new DataFormatter();
            this.df = new DecimalFormat("#.00");

        } catch (FileNotFoundException ex) {
            Logger.getLogger(LectorXLS_Link.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "El fichero no funciona");
        } catch (IOException ex) {
            Logger.getLogger(LectorXLS_Link.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "No se puede creer el fichero");
        } catch (EncryptedDocumentException ex) {
            Logger.getLogger(LectorXLS_Link.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Documento encriptado");
        } catch (InvalidFormatException ex) {
            Logger.getLogger(LectorXLS_Link.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    // Lee la tabla importada y llena el Jtable del formulario
    public void LeerExcel(JTable tabla) {

        DefaultTableModel dt = new DefaultTableModel() {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        Iterator i = this.hoja.iterator();
        Object[] fila;
        int fil = 0;

        while (i.hasNext()) {
            Row filaSiguiente = (Row) i.next();
            Iterator iCelda = filaSiguiente.cellIterator();

            // contamos el numero de celdas de la fila
            int largoFila = filaSiguiente.getLastCellNum();
            fila = new Object[largoFila];
            int j = 0;
            if (fil != 0) {
                while (iCelda.hasNext()) {

                    Cell celda = (Cell) iCelda.next();
                    String contenidoCelda = formatoDeDatos.formatCellValue(celda);

                    if (contenidoCelda == null) {
                        contenidoCelda = "null";
                    }

                    fila[j] = contenidoCelda;
                    j++;

                }
            } else {
                while (iCelda.hasNext()) { // Si la fila es la numero 0 osea la primera se ponen los encabezados
                    Cell celda = (Cell) iCelda.next();
                    String contenidoCelda = formatoDeDatos.formatCellValue(celda);
                    dt.addColumn(contenidoCelda);
                }

            }
            if (fil != 0) {
                dt.addRow(fila);
            } else {
                fil = 1;
            }

        }
        tabla.setModel(dt);
    }

    public void leerTabla(JTable tabla, String path) {

        this.escritor = new EscritorTxt(path + ".txt");

        String IdOk = null, concepto = null, fechaPrimerVto = null, fechaSegundoVto = null, montoPrimerVto = null, montoSegundoVto = null,
                numeroFactura = null, documentoOrigen = null, conceptoAnterior = null, idDeuda = null;
        int cont = 1;

        String fechaStr = DateTimeFormatter.ofPattern("MMyy").format(LocalDateTime.now());
        String fechaPrimerFila = DateTimeFormatter.ofPattern("yyMMdd").format(LocalDateTime.now());

        String primerFila = "HRFACTURACIONGK0" + fechaPrimerFila + "00001";
        primerFila = primerFila + asignarEspacios(primerFila);
        
        this.escritor.escribir(primerFila, true);

        for (int j = 0; j < tabla.getRowCount(); j++) {
            this.contadorFilas++;
            for (int i = 0; i < tabla.getColumnCount(); i++) {

                String valorCelda = (String) tabla.getValueAt(j, i);

                // Creo el ID
                if (i == 0) {
                    char[] idArray = valorCelda.toCharArray();
                    char id[] = {'0', '0', '0', '0', '0', '0', '0', '0'};
                    int y = id.length - idArray.length;

                    for (char c : idArray) {
                        id[y] = c;
                        y++;
                    }
                    IdOk = valueOf(id);
                }

                // Tomo la informacion de la columna con el documento de origen
                if (i == 1) {

                    documentoOrigen = valorCelda;

                    if (valorCelda.contains("SUB")) {
                        concepto = "001"; // Facturas de abono

                    } else if (valorCelda.contains("SO") && tabla.getValueAt(j, i + 4).equals("3500")) {
                        concepto = "003"; // Facturas Instalacion

                    } else if (valorCelda.contains("SO")) {
                        concepto = "002"; // Facturas de proporcional

                    } else {
                        concepto = "004"; // Otros conceptos

                    }
                }

                // Se comprueba si el cliente tiene varias deudas del mismo concepto para eso se compara el id
                // y los conceptos que tiene anteriormente
                if (i == 2) {
                    String idAnterior = null;
                    String IdActual = null;
                    if (j != 0) {
                        IdActual = (String) tabla.getValueAt(j, 0);
                        idAnterior = (String) tabla.getValueAt(j - 1, 0);
                        if (idAnterior.equals(IdActual) && concepto.equals(conceptoAnterior)) {
                            idDeuda = cont + fechaStr;
                            cont++;
                        } else {
                            idDeuda = "0" + fechaStr;
                            cont = 1;
                        }
                    } else {
                        idDeuda = "0" + fechaStr;
                    }

                }

                // Establezco las fechas de vencimiento
                if (i == 3) {
                    fechaPrimerVto = valorCelda.replace("-", "");
                    fechaPrimerVto = fechaPrimerVto.substring(2, 8);
                    fechaSegundoVto = fechaPrimerVto.substring(0, 4) + "25";
                }

                //Establezco los montos del primer vencimiento
                if (i == 4) {

                    String totalAdeudado = (String) tabla.getValueAt(j, i + 3);
                    totalAdeudado = totalAdeudado.replace(",", ".");

                    String totalConRecargo = (String) tabla.getValueAt(j, i + 1);
                    totalConRecargo = totalConRecargo.replace(",", ".");

                    String totalSinRecargo = valorCelda.replace(",", ".");

                    if (totalSinRecargo.equals(totalAdeudado)) {

                        this.montoTotalPrimerVto += Double.parseDouble(totalSinRecargo);

                        montoPrimerVto = totalSinRecargo.replace(".", ""); // si se debe el total de la factura y el adeudado es igual 
                        char[] montoArray = montoPrimerVto.toCharArray();
                        char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
                        int y = monto.length - montoArray.length;

                        for (char c : montoArray) {
                            monto[y] = c;
                            y++;
                        }
                        montoPrimerVto = valueOf(monto);

                    } else {

                        if (!totalAdeudado.equals(totalConRecargo)) {
                            this.montoTotalPrimerVto += Double.parseDouble(totalAdeudado);

                            montoPrimerVto = totalAdeudado;
                            montoPrimerVto = montoPrimerVto.replace(".", ""); // Si la factura esta abierta pero no se debe completa
                            char[] montoArray = montoPrimerVto.toCharArray();
                            char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
                            int y = monto.length - montoArray.length;

                            for (char c : montoArray) {
                                monto[y] = c;
                                y++;
                            }
                            montoPrimerVto = valueOf(monto);
                        } else {
                            this.montoTotalPrimerVto += Double.parseDouble(totalSinRecargo);
                            montoPrimerVto = totalSinRecargo.replace(".", ""); // si se debe el total de la factura y el adeudado es igual 
                            char[] montoArray = montoPrimerVto.toCharArray();
                            char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
                            int y = monto.length - montoArray.length;
                            for (char c : montoArray) {
                                monto[y] = c;
                                y++;
                            }
                            montoPrimerVto = valueOf(monto);
                        }

                    }
                }

                // Establezco los montos del segundo vencimiento
                if (i == 5) {

                    if (!valorCelda.equals("0,00")) {
                        String totalAdeudado = (String) tabla.getValueAt(j, i + 2);
                        totalAdeudado = totalAdeudado.replace(",", ".");
                        valorCelda = valorCelda.replace(",", ".");

                        if (valorCelda.equals(totalAdeudado)) {
                            valorCelda = valorCelda.replace(",", ".");
                            this.montoTotalSegundoVto += Double.parseDouble(valorCelda);

                            montoSegundoVto = valorCelda.replace(".", ""); // Si es distinto de 0 el primer vencimiento tiene recargo
                            char[] montoArray = montoSegundoVto.toCharArray();
                            char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
                            int y = monto.length - montoArray.length;

                            for (char c : montoArray) {
                                monto[y] = c;
                                y++;
                            }
                            montoSegundoVto = valueOf(monto);
                        } else {
                            this.montoTotalSegundoVto += Double.parseDouble(valorCelda);

                            montoSegundoVto = valorCelda;
                            montoSegundoVto = montoSegundoVto.replace(".", ""); // Si la factura esta abierta pero no se debe completa
                            char[] montoArray = montoSegundoVto.toCharArray();
                            char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
                            int y = monto.length - montoArray.length;

                            for (char c : montoArray) {
                                monto[y] = c;
                                y++;
                            }
                            montoSegundoVto = valueOf(monto);

                        }

                    } else {
                        String monto1 = (String) tabla.getValueAt(j, i + 2);
                        monto1 = monto1.replace(",", ".");
                        this.montoTotalSegundoVto += Double.parseDouble(monto1);
                        montoSegundoVto = montoPrimerVto; // si es igual a 0 el primer vencimiento no tiene re cargo y es igual al monto del primer vencimiento
                    }

                }

                // Obtengo el numero de factura
                if (i == 6) {
                    numeroFactura = valorCelda.replace("-", " ");
                    numeroFactura = numeroFactura.substring(3);
                }

            }

            char[] charArray = new char[48];

            for (int i = 0; i < charArray.length; i++) {
                charArray[i] = '0';

            }

            String fila = idDeuda + concepto + IdOk + "           " + fechaPrimerVto + montoPrimerVto + fechaSegundoVto + montoSegundoVto + "000000000000000000"
                    + numeroFactura + " " + documentoOrigen;

            fila = fila + asignarEspacios(fila);

            this.escritor.escribir(fila, true);

            conceptoAnterior = concepto;

        }

        String ultimaFila = "TRFACTURACION" + cerosIzquierda(String.valueOf(this.contadorFilas), 8)
                + cerosIzquierda(String.valueOf(df.format(montoTotalPrimerVto)).replace(",", ""), 18)
                + cerosIzquierda(String.valueOf(df.format(montoTotalSegundoVto)).replace(",", ""), 18) + "000000000000000000";

        ultimaFila = ultimaFila + asignarEspacios(ultimaFila);
        
        this.escritor.escribir(ultimaFila, true);

        String ultimaFecha = (String) tabla.getValueAt(0, 3);
        ultimaFecha = ultimaFecha.replace("-", "");
        archivoDeControl(path, this.contadorFilas, this.montoTotalPrimerVto, this.montoTotalSegundoVto, ultimaFecha);

        JOptionPane.showMessageDialog(null, "ARCHIVOS GENERADOS");
        escritor.abrirarchivo();
    }

    public String asignarEspacios(String fila) {
        String espacios = "";
        for (int b = fila.length(); b <= 130; b++) {
            espacios += " ";
        }
        return espacios;
    }

    public String asignarEspaciosControl(String fila) {
        String espacios = "";
        for (int b = fila.length(); b <= 74; b++) {
            espacios += " ";
        }
        return espacios;
    }

    public String cerosIzquierda(String fila, int pos) {
        String filaRet = null;
        char[] arrayFila = fila.toCharArray();
        char[] filaArray = new char[pos];

        for (int i = 0; i < filaArray.length; i++) {
            filaArray[i] = '0';
        }

        int y = pos - arrayFila.length;

        for (char c : arrayFila) {
            filaArray[y] = c;
            y++;
        }
        filaRet = valueOf(filaArray);
        return filaRet;
    }

    public void archivoDeControl(String path, int registros, double montoPrimVto, double montoSeguVto, String ultimaFecha) {

        String nombreArchivo = path.substring(path.length() - 8);
        String nombreArchivoControl = nombreArchivo.replace("P", "C");
        path = path.substring(0, path.length() - 8) + nombreArchivoControl + ".txt";

        this.escritorControl = new EscritorTxt(path);

        String fechaStr = DateTimeFormatter.ofPattern("yyyyMMdd").format(LocalDateTime.now());

        int registrosControl = registros * 133;
        String registrosStr = valueOf(registrosControl);
        registrosStr = cerosIzquierda(registrosStr, 10);

        String fila1 = "HRPASCTRL" + fechaStr + "GK0" + nombreArchivo + registrosStr;
        fila1 = fila1 + asignarEspaciosControl(fila1);

        String fila2 = "LOTES" + "00001" + cerosIzquierda(valueOf(registros), 8)
                + cerosIzquierda(String.valueOf(df.format(montoTotalPrimerVto)).replace(",", ""), 18)
                + cerosIzquierda(String.valueOf(df.format(montoTotalSegundoVto)).replace(",", ""), 18) + "000000000000000000";

        fila2 = fila2 + asignarEspaciosControl(fila2);

        String fila3 = "FINAL" + cerosIzquierda(valueOf(registros), 8) + cerosIzquierda(String.valueOf(df.format(montoTotalPrimerVto)).replace(",", ""), 18)
                + cerosIzquierda(String.valueOf(df.format(montoTotalSegundoVto)).replace(",", ""), 18) + "000000000000000000" + ultimaFecha;

        this.escritorControl.escribir(fila1, true);
        this.escritorControl.escribir(fila2, true);
        this.escritorControl.escribir(fila3, true);
    }
}
