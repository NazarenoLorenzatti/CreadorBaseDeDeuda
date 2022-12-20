package UI;

import FileReader.LectorXLS_PMC;
import java.io.File;
import java.text.ParseException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.filechooser.FileNameExtensionFilter;

public class FrmPMC_Excel extends javax.swing.JFrame {

    private LectorXLS_PMC leerXLS;

    public FrmPMC_Excel() {
        initComponents();

    }

    public String abrirArchivos() {
        String directorio = null;
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos Excel/texto", "xls", "xlsx", "txt");
        fc.setFileFilter(filter);

        int seleccion = fc.showOpenDialog(this);

        if (seleccion == JFileChooser.APPROVE_OPTION) {

            File fichero = fc.getSelectedFile();
            directorio = fichero.getAbsolutePath();
        }

        return directorio;

    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        BtnHome = new javax.swing.JButton();
        Importar = new javax.swing.JButton();
        jBtnGenerarTXT = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(33, 46, 49));

        BtnHome.setIcon(new javax.swing.ImageIcon(getClass().getResource("/UI/Images/icons8-casa-50.png"))); // NOI18N
        BtnHome.setBorderPainted(false);
        BtnHome.setContentAreaFilled(false);
        BtnHome.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        BtnHome.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/UI/Images/icons8-casa-50 - roll.png"))); // NOI18N
        BtnHome.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                BtnHomeActionPerformed(evt);
            }
        });

        Importar.setBackground(new java.awt.Color(0, 47, 42));
        Importar.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        Importar.setForeground(new java.awt.Color(255, 255, 255));
        Importar.setText("Importar Tabla");
        Importar.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        Importar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ImportarActionPerformed(evt);
            }
        });

        jBtnGenerarTXT.setBackground(new java.awt.Color(0, 47, 42));
        jBtnGenerarTXT.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jBtnGenerarTXT.setForeground(new java.awt.Color(255, 255, 255));
        jBtnGenerarTXT.setText("Generar TXT");
        jBtnGenerarTXT.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jBtnGenerarTXT.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBtnGenerarTXTActionPerformed(evt);
            }
        });

        jTable1.setBackground(new java.awt.Color(255, 255, 255));
        jTable1.setForeground(new java.awt.Color(0, 0, 0));
        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {},
                {},
                {},
                {}
            },
            new String [] {

            }
        ));
        jScrollPane1.setViewportView(jTable1);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(234, 234, 234)
                        .addComponent(Importar, javax.swing.GroupLayout.PREFERRED_SIZE, 194, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(163, 163, 163)
                        .addComponent(jBtnGenerarTXT, javax.swing.GroupLayout.PREFERRED_SIZE, 206, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(24, 24, 24)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(BtnHome)
                            .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 1005, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addContainerGap(30, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(22, 22, 22)
                .addComponent(BtnHome, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 22, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jBtnGenerarTXT, javax.swing.GroupLayout.PREFERRED_SIZE, 77, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(Importar, javax.swing.GroupLayout.PREFERRED_SIZE, 77, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(65, 65, 65)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 336, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(28, 28, 28))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void ImportarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ImportarActionPerformed
        leerXLS = new LectorXLS_PMC(abrirArchivos());
        leerXLS.LeerExcel(jTable1);
    }//GEN-LAST:event_ImportarActionPerformed

    private void jBtnGenerarTXTActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBtnGenerarTXTActionPerformed

        try {
            leerXLS.leerTabla(jTable1, abrirArchivos());
        } catch (ParseException ex) {
            Logger.getLogger(FrmPMC_Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_jBtnGenerarTXTActionPerformed

    private void BtnHomeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_BtnHomeActionPerformed
        FrmHome f = new FrmHome();
        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        f.setLocationRelativeTo(null);
        f.setVisible(true);
        f.show();
        dispose();
    }//GEN-LAST:event_BtnHomeActionPerformed

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton BtnHome;
    private javax.swing.JButton Importar;
    private javax.swing.JButton jBtnGenerarTXT;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
}
