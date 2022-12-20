package FileReader;

import FileWriter.EscritorTxt;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import static java.lang.String.valueOf;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.time.LocalDate;
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

public class LectorXLS_PMC {

    private String path;
    private FileInputStream inputStream;
    private Workbook wb;
    private Sheet hoja;
    private EscritorTxt escritor;
    private DataFormatter formatoDeDatos;
    private int contadorFilas = 0;
    private double montoTotalPrimerVto = 0;
    private DecimalFormat df;
    private String mensajeTicket;
    private String mensajePantalla;

    public LectorXLS_PMC(String path) {

        try {

            this.path = path;
            this.inputStream = new FileInputStream(new File(path));
            this.wb = WorkbookFactory.create(inputStream);
            this.hoja = wb.getSheetAt(0);
            this.formatoDeDatos = new DataFormatter();
            this.df = new DecimalFormat("#.00");
            

        } catch (FileNotFoundException ex) {
            Logger.getLogger(LectorXLS_PMC.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "El fichero no funciona");
        } catch (IOException ex) {
            Logger.getLogger(LectorXLS_PMC.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "No se puede creer el fichero");
        } catch (EncryptedDocumentException ex) {
            Logger.getLogger(LectorXLS_PMC.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Documento encriptado");
        } catch (InvalidFormatException ex) {
            Logger.getLogger(LectorXLS_PMC.class.getName()).log(Level.SEVERE, null, ex);
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

    public void leerTabla(JTable tabla, String path) throws ParseException {        
        
        String IdOk = null, fechaPrimerVto = null, fechaSegundoVto = null, montoPrimerVto = null, montoSegundoVto = null,
                numeroFactura = null, dniCliente = null, fechaTercerVto = null;
        int cont = 1;
        
        this.escritor = new EscritorTxt(path);

        String fechaPrimerFila = DateTimeFormatter.ofPattern("yyyyMMdd").format(LocalDateTime.now());
        
        String primerFila = "0400SBHN" + fechaPrimerFila + cerosDerecha("", 264);
        this.escritor.escribir(primerFila, true);

        for (int j = 0; j < tabla.getRowCount(); j++) {
            this.contadorFilas++;
            for (int i = 0; i < tabla.getColumnCount(); i++) {

                String valorCelda = (String) tabla.getValueAt(j, i);

                // Creo el ID
                if (i == 0) {
                    IdOk = valorCelda;

                    if (j != 0) {
                        String idAnterior = (String) tabla.getValueAt(j - 1, i);
                        if (idAnterior.equals(valorCelda)) {
                            cont++;
                            IdOk += "-" + cont;
                        } else {
                            cont = 1;
                            IdOk += "-" + cont;
                        }
                    } else {
                        IdOk += "-" + cont;
                    }

                    IdOk = asignarEspacios(IdOk, 20);

                }

                // Tomo la informacion de la columna con el Dni, informacion por medio de la cual el cliente va a buscar su deuda
                if (i == 1) {
                    dniCliente = "5" + asignarEspacios(valorCelda, 19);
                }

                // Establecemos las fechas de vencimiento               
                if (i == 2) {

                    LocalDate fecha = LocalDate.parse(valorCelda);

                    int mes = fecha.getMonthValue();

                    asignarMes(mes);

                    fechaPrimerVto = fecha.toString();
                    fechaPrimerVto = fechaPrimerVto.replace("-", "");

                    fechaSegundoVto = fecha.plusDays(15).toString();
                    fechaSegundoVto = fechaSegundoVto.replace("-", "");

                    fechaTercerVto = fecha.plusDays(75).toString();
                    fechaTercerVto = fechaTercerVto.replace("-", "");
                }

                //Establezco los montos del primer vencimiento
                if (i == 3) {

                    String totalAdeudado = (String) tabla.getValueAt(j, i + 3);
                    totalAdeudado = totalAdeudado.replace(",", ".");

                    String totalConRecargo = (String) tabla.getValueAt(j, i + 1);
                    totalConRecargo = totalConRecargo.replace(",", ".");

                    String totalSinRecargo = valorCelda.replace(",", ".");

                    if (totalSinRecargo.equals(totalAdeudado)) {

                        this.montoTotalPrimerVto += Double.parseDouble(totalSinRecargo);

                        montoPrimerVto = totalSinRecargo.replace(".", ""); // si se debe el total de la factura y el adeudado es igual 
                        char[] montoArray = montoPrimerVto.toCharArray();
                        char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
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
                            char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
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
                            char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
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
                if (i == 4) {

                    if (!valorCelda.equals("0,00")) {
                        String totalAdeudado = (String) tabla.getValueAt(j, i + 2);
                        totalAdeudado = totalAdeudado.replace(",", ".");
                        valorCelda = valorCelda.replace(",", ".");

                        if (valorCelda.equals(totalAdeudado)) {
                            valorCelda = valorCelda.replace(",", ".");

                            montoSegundoVto = valorCelda.replace(".", ""); // Si es distinto de 0 el primer vencimiento tiene recargo
                            char[] montoArray = montoSegundoVto.toCharArray();
                            char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
                            int y = monto.length - montoArray.length;

                            for (char c : montoArray) {
                                monto[y] = c;
                                y++;
                            }
                            montoSegundoVto = valueOf(monto);
                        } else {

                            montoSegundoVto = valorCelda;
                            montoSegundoVto = montoSegundoVto.replace(".", ""); // Si la factura esta abierta pero no se debe completa
                            char[] montoArray = montoSegundoVto.toCharArray();
                            char monto[] = {'0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0'};
                            int y = monto.length - montoArray.length;

                            for (char c : montoArray) {
                                monto[y] = c;
                                y++;
                            }
                            montoSegundoVto = valueOf(monto);

                        }

                    } else {
                        montoSegundoVto = montoPrimerVto; // si es igual a 0 el primer vencimiento no tiene re cargo y es igual al monto del primer vencimiento
                    }

                }

                // Obtengo el numero de factura
                if (i == 5) {
                    numeroFactura = valorCelda.replaceFirst("-", " ");
                    numeroFactura = numeroFactura.substring(3);
                }

            }

            String fila = dniCliente + IdOk + "0" + fechaPrimerVto + montoPrimerVto + fechaSegundoVto + montoSegundoVto + fechaTercerVto + montoSegundoVto
                    + "000000000000000000" + dniCliente.substring(1) +" " + asignarEspacios(mensajeTicket, 40) + asignarEspacios(mensajePantalla, 15)
                    + asignarEspacios(numeroFactura, 60) + cerosDerecha("", 29);

            this.escritor.escribir(fila, true);

        }
        
        String monto = String.valueOf(df.format(montoTotalPrimerVto));
        
        String ultimaFila = "9400SBHN" + fechaPrimerFila + cerosIzquierda(valueOf(this.contadorFilas), 7)+ "0000000" 
                + cerosIzquierda(String.valueOf(df.format(montoTotalPrimerVto)).replace(",", ""), 16)
                + cerosDerecha("", 234);
        
        this.escritor.escribir(ultimaFila, true);

        JOptionPane.showMessageDialog(null, "ARCHIVOS GENERADOS");
        escritor.abrirarchivo();
    }
    

    public String asignarEspacios(String fila, int pos) {
        String filaRet = null;
        char[] arrayInput = fila.toCharArray();
        char[] arrayOutput = new char[pos];

        for (int i = 0; i < arrayOutput.length; i++) {
            arrayOutput[i] = ' ';
        }

        for (int i = 0; i < arrayInput.length; i++) {
            arrayOutput[i] = arrayInput[i];
        }

        filaRet = valueOf(arrayOutput);
        return filaRet;
    }
    

    public String cerosDerecha(String fila, int pos) {
        String filaRet = null;
        char[] arrayInput = fila.toCharArray();
        char[] arrayOutput = new char[pos];

        for (int i = 0; i < arrayOutput.length; i++) {
            arrayOutput[i] = '0';
        }

        for (int i = 0; i < arrayInput.length; i++) {
            arrayOutput[i] = arrayInput[i];
        }

        filaRet = valueOf(arrayOutput);
        return filaRet;
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
    

    public void asignarMes(int mes) {
        switch (mes) {
            case 1:
                mensajeTicket = "FAC ENERO ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC ENE";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 2:
                mensajeTicket = "FAC FEBRERO ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC FEB";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 3:
                mensajeTicket = "FAC MARZO ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC MAR";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 4:
                mensajeTicket = "FAC ABRIL ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC ABR";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 5:
                mensajeTicket = "FAC MAYO ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC MAY";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 6:
                mensajeTicket = "FAC JUNIO ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC JUN";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 7:
                mensajeTicket = "FAC JULIO ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC JUL";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 8:
                mensajeTicket = "FAC AGOSTO ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC AGO";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 9:
                mensajeTicket = "FAC SEPTIEMBRE ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC SEP";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 10:
                mensajeTicket = "FAC OCTUBRE ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC OCT";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 11:
                mensajeTicket = "FAC NOVIEMBRE ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC NOV";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
            case 12:
                mensajeTicket = "FAC DICIEMBRE ULTRAFIBRA";
                mensajeTicket = asignarEspacios(mensajeTicket, 40);

                mensajePantalla = "FAC DIC";
                mensajePantalla = asignarEspacios(mensajePantalla, 15);
                break;
        }
    }

}
