package surat;

import fungsi.WarnaTable;
import fungsi.akses;
import fungsi.batasInput;
import fungsi.koneksiDB;
import fungsi.sekuel;
import fungsi.validasi;
import java.awt.BorderLayout;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Frame;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.RejectedExecutionException;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.SwingUtilities;
import javax.swing.WindowConstants;
import javax.swing.event.DocumentEvent;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import kepegawaian.DlgCariPetugas;
import net.sf.jasperreports.engine.DefaultJasperReportsContext;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.design.JRCompiler;
import net.sf.jasperreports.engine.design.JRJavacCompiler;
import net.sf.jasperreports.engine.SimpleJasperReportsContext;

public final class SuratPernyataanEdukasiRujukan extends javax.swing.JDialog {
    private final DefaultTableModel tabMode;
    private final Connection koneksi = koneksiDB.condb();
    private final sekuel Sequel = new sekuel();
    private final validasi Valid = new validasi();
    private PreparedStatement ps;
    private ResultSet rs;
    private DlgCariPetugas petugas;
    private final ExecutorService executor = Executors.newSingleThreadExecutor();
    private volatile boolean ceksukses = false;
    private final SimpleDateFormat jamFormat = new SimpleDateFormat("HH:mm");
    private widget.TextBox AlamatDirujuk;
    private widget.TextBox AlamatPenyetuju;
    private widget.Button BtnAll;
    private widget.Button BtnBatal;
    private widget.Button BtnCari;
    private widget.Button BtnEdit;
    private widget.Button BtnHapus;
    private widget.Button BtnKeluar;
    private widget.Button BtnPetugas;
    private widget.Button BtnPrint;
    private widget.Button BtnSimpan;
    private widget.CekBox ChkInput;
    private widget.Tanggal DTPCari1;
    private widget.Tanggal DTPCari2;
    private widget.PanelBiasa FormInput;
    private widget.ComboBox Hubungan;
    private widget.TextBox JK;
    private widget.ComboBox JKDirujuk;
    private widget.ComboBox JKPenyetuju;
    private widget.TextBox Jam;
    private widget.TextBox KdPetugas;
    private widget.Label LCount;
    private widget.TextBox NamaDirujuk;
    private widget.TextBox NamaPenyetuju;
    private widget.TextBox NmPetugas;
    private widget.TextBox NoSurat;
    private javax.swing.JPanel PanelInput;
    private widget.TextBox Saksi1;
    private widget.TextBox Saksi2;
    private widget.ScrollPane Scroll;
    private widget.TextBox TCari;
    private widget.TextBox TNoRM;
    private widget.TextBox TNoRw;
    private widget.TextBox TPasien;
    private widget.Tanggal Tanggal;
    private widget.TextBox Umur;
    private widget.TextBox UmurDirujuk;
    private widget.TextBox UmurPenyetuju;
    private widget.InternalFrame internalFrame1;
    private javax.swing.JPanel jPanel3;
    private widget.Label jLabel10;
    private widget.Label jLabel11;
    private widget.Label jLabel12;
    private widget.Label jLabel13;
    private widget.Label jLabel14;
    private widget.Label jLabel15;
    private widget.Label jLabel16;
    private widget.Label jLabel17;
    private widget.Label jLabel18;
    private widget.Label jLabel19;
    private widget.Label jLabel20;
    private widget.Label jLabel21;
    private widget.Label jLabel22;
    private widget.Label jLabel23;
    private widget.Label jLabel24;
    private widget.Label jLabel25;
    private widget.Label jLabel26;
    private widget.Label jLabel27;
    private widget.Label jLabel28;
    private widget.Label jLabel29;
    private widget.Label jLabel3;
    private widget.Label jLabel30;
    private widget.Label jLabel4;
    private widget.Label jLabel5;
    private widget.Label jLabel6;
    private widget.Label jLabel7;
    private widget.Label jLabel8;
    private widget.Label jLabel9;
    private widget.panelisi panelGlass8;
    private widget.panelisi panelGlass9;
    private widget.Table tbObat;

    public SuratPernyataanEdukasiRujukan(Frame parent, boolean modal) {
        super(parent, modal);
        cekTabel();
        initComponents();
        this.setLocation(8, 1);
        setSize(628, 650);

        tabMode = new DefaultTableModel(null, new Object[]{
            "No.Surat", "No.Rawat", "No.R.M.", "Nama Pasien", "Umur Pasien", "J.K. Pasien",
            "Tanggal", "Jam", "Nama Penyetuju", "Alamat Penyetuju", "Umur Penyetuju",
            "J.K. Penyetuju", "Hubungan", "Nama Dirujuk", "Alamat Dirujuk", "Umur Dirujuk",
            "J.K. Dirujuk", "Saksi 1", "Saksi 2", "NIP", "Nama Petugas"
        }) {
            @Override
            public boolean isCellEditable(int rowIndex, int colIndex) {
                return false;
            }
        };
        tbObat.setModel(tabMode);
        tbObat.setPreferredScrollableViewportSize(new Dimension(500, 500));
        tbObat.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        for (int i = 0; i < 21; i++) {
            TableColumn column = tbObat.getColumnModel().getColumn(i);
            if (i == 0 || i == 1) {
                column.setPreferredWidth(105);
            } else if (i == 2) {
                column.setPreferredWidth(70);
            } else if (i == 3 || i == 8 || i == 13 || i == 20) {
                column.setPreferredWidth(150);
            } else if (i == 4 || i == 10 || i == 15) {
                column.setPreferredWidth(80);
            } else if (i == 5 || i == 11 || i == 16) {
                column.setPreferredWidth(70);
            } else if (i == 6) {
                column.setPreferredWidth(75);
            } else if (i == 7) {
                column.setPreferredWidth(55);
            } else if (i == 9 || i == 14) {
                column.setPreferredWidth(220);
            } else if (i == 12) {
                column.setPreferredWidth(120);
            } else if (i == 17 || i == 18) {
                column.setPreferredWidth(120);
            } else if (i == 19) {
                column.setPreferredWidth(90);
            }
        }
        tbObat.setDefaultRenderer(Object.class, new WarnaTable());
        tbObat.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                getData();
            }
        });
        tbObat.addKeyListener(new java.awt.event.KeyAdapter() {
            @Override
            public void keyReleased(KeyEvent e) {
                getData();
            }
        });

        TNoRw.setDocument(new batasInput((byte) 17).getKata(TNoRw));
        NoSurat.setDocument(new batasInput((byte) 20).getKata(NoSurat));
        KdPetugas.setDocument(new batasInput((byte) 20).getKata(KdPetugas));
        TCari.setDocument(new batasInput(100).getKata(TCari));
        NamaPenyetuju.setDocument(new batasInput((byte) 50).getKata(NamaPenyetuju));
        AlamatPenyetuju.setDocument(new batasInput(150).getKata(AlamatPenyetuju));
        UmurPenyetuju.setDocument(new batasInput((byte) 20).getKata(UmurPenyetuju));
        Jam.setDocument(new batasInput((byte) 5).getKata(Jam));
        Saksi1.setDocument(new batasInput((byte) 50).getKata(Saksi1));
        Saksi2.setDocument(new batasInput((byte) 50).getKata(Saksi2));

        TNoRM.setEditable(false);
        TPasien.setEditable(false);
        NamaDirujuk.setEditable(false);
        AlamatDirujuk.setEditable(false);
        UmurDirujuk.setEditable(false);
        JKDirujuk.setEnabled(false);
        ChkInput.setVisible(false);
        ChkInput.setSelected(true);
        isForm();
        emptTeks();
    }

    @SuppressWarnings("unchecked")
    private void initComponents() {
        internalFrame1 = new widget.InternalFrame();
        Scroll = new widget.ScrollPane();
        tbObat = new widget.Table();
        jPanel3 = new javax.swing.JPanel();
        panelGlass8 = new widget.panelisi();
        BtnSimpan = new widget.Button();
        BtnBatal = new widget.Button();
        BtnHapus = new widget.Button();
        BtnEdit = new widget.Button();
        BtnPrint = new widget.Button();
        BtnAll = new widget.Button();
        BtnKeluar = new widget.Button();
        panelGlass9 = new widget.panelisi();
        jLabel19 = new widget.Label();
        DTPCari1 = new widget.Tanggal();
        jLabel21 = new widget.Label();
        DTPCari2 = new widget.Tanggal();
        jLabel6 = new widget.Label();
        TCari = new widget.TextBox();
        BtnCari = new widget.Button();
        jLabel7 = new widget.Label();
        LCount = new widget.Label();
        PanelInput = new javax.swing.JPanel();
        FormInput = new widget.PanelBiasa();
        jLabel4 = new widget.Label();
        TNoRw = new widget.TextBox();
        TPasien = new widget.TextBox();
        TNoRM = new widget.TextBox();
        jLabel3 = new widget.Label();
        NoSurat = new widget.TextBox();
        jLabel16 = new widget.Label();
        Tanggal = new widget.Tanggal();
        jLabel14 = new widget.Label();
        Jam = new widget.TextBox();
        jLabel10 = new widget.Label();
        Hubungan = new widget.ComboBox();
        jLabel22 = new widget.Label();
        jLabel8 = new widget.Label();
        NamaPenyetuju = new widget.TextBox();
        jLabel9 = new widget.Label();
        AlamatPenyetuju = new widget.TextBox();
        jLabel11 = new widget.Label();
        UmurPenyetuju = new widget.TextBox();
        jLabel12 = new widget.Label();
        JKPenyetuju = new widget.ComboBox();
        jLabel23 = new widget.Label();
        jLabel24 = new widget.Label();
        NamaDirujuk = new widget.TextBox();
        jLabel25 = new widget.Label();
        AlamatDirujuk = new widget.TextBox();
        jLabel26 = new widget.Label();
        UmurDirujuk = new widget.TextBox();
        jLabel27 = new widget.Label();
        JKDirujuk = new widget.ComboBox();
        jLabel28 = new widget.Label();
        Saksi1 = new widget.TextBox();
        jLabel29 = new widget.Label();
        Saksi2 = new widget.TextBox();
        jLabel17 = new widget.Label();
        KdPetugas = new widget.TextBox();
        NmPetugas = new widget.TextBox();
        BtnPetugas = new widget.Button();
        ChkInput = new widget.CekBox();
        JK = new widget.TextBox();
        Umur = new widget.TextBox();
        jLabel5 = new widget.Label();
        jLabel13 = new widget.Label();
        jLabel15 = new widget.Label();
        jLabel18 = new widget.Label();
        jLabel20 = new widget.Label();
        jLabel30 = new widget.Label();

        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        setUndecorated(true);
        setResizable(false);
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowOpened(WindowEvent evt) {
                formWindowOpened(evt);
            }
        });

        internalFrame1.setBorder(javax.swing.BorderFactory.createTitledBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(240, 245, 235)), "::[ Surat Pernyataan Edukasi Rujukan ]::", javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Tahoma", 0, 11), new java.awt.Color(50, 50, 50)));
        internalFrame1.setLayout(new BorderLayout(1, 1));

        Scroll.setOpaque(true);
        Scroll.setPreferredSize(new Dimension(452, 200));
        tbObat.setAutoCreateRowSorter(true);
        tbObat.setToolTipText("Silahkan klik untuk memilih data yang mau diedit ataupun dihapus");
        Scroll.setViewportView(tbObat);
        internalFrame1.add(Scroll, BorderLayout.CENTER);

        jPanel3.setLayout(new BorderLayout(1, 1));

        panelGlass8.setPreferredSize(new Dimension(55, 53));
        panelGlass8.setLayout(new FlowLayout(FlowLayout.LEFT, 5, 9));

        BtnSimpan.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/save-16x16.png")));
        BtnSimpan.setMnemonic('S');
        BtnSimpan.setText("Simpan");
        BtnSimpan.setToolTipText("Alt+S");
        BtnSimpan.setPreferredSize(new Dimension(100, 30));
        BtnSimpan.addActionListener(evt -> BtnSimpanActionPerformed(evt));
        panelGlass8.add(BtnSimpan);

        BtnBatal.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/stop_f2.png")));
        BtnBatal.setMnemonic('B');
        BtnBatal.setText("Baru");
        BtnBatal.setToolTipText("Alt+B");
        BtnBatal.setPreferredSize(new Dimension(100, 30));
        BtnBatal.addActionListener(evt -> BtnBatalActionPerformed(evt));
        panelGlass8.add(BtnBatal);

        BtnHapus.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/stop_f2.png")));
        BtnHapus.setMnemonic('H');
        BtnHapus.setText("Hapus");
        BtnHapus.setToolTipText("Alt+H");
        BtnHapus.setPreferredSize(new Dimension(100, 30));
        BtnHapus.addActionListener(evt -> BtnHapusActionPerformed(evt));
        panelGlass8.add(BtnHapus);

        BtnEdit.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/accept.png")));
        BtnEdit.setMnemonic('E');
        BtnEdit.setText("Ganti");
        BtnEdit.setToolTipText("Alt+E");
        BtnEdit.setPreferredSize(new Dimension(100, 30));
        BtnEdit.addActionListener(evt -> BtnEditActionPerformed(evt));
        panelGlass8.add(BtnEdit);

        BtnPrint.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/b_print.png")));
        BtnPrint.setMnemonic('T');
        BtnPrint.setText("Cetak Surat");
        BtnPrint.setToolTipText("Alt+T");
        BtnPrint.setPreferredSize(new Dimension(110, 30));
        BtnPrint.addActionListener(evt -> BtnPrintActionPerformed(evt));
        panelGlass8.add(BtnPrint);

        BtnAll.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Search-16x16.png")));
        BtnAll.setMnemonic('M');
        BtnAll.setText("Semua");
        BtnAll.setToolTipText("Alt+M");
        BtnAll.setPreferredSize(new Dimension(100, 30));
        BtnAll.addActionListener(evt -> BtnAllActionPerformed(evt));
        panelGlass8.add(BtnAll);

        BtnKeluar.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/exit.png")));
        BtnKeluar.setMnemonic('K');
        BtnKeluar.setText("Keluar");
        BtnKeluar.setToolTipText("Alt+K");
        BtnKeluar.setPreferredSize(new Dimension(100, 30));
        BtnKeluar.addActionListener(evt -> BtnKeluarActionPerformed(evt));
        panelGlass8.add(BtnKeluar);

        jPanel3.add(panelGlass8, BorderLayout.PAGE_END);

        panelGlass9.setPreferredSize(new Dimension(55, 43));
        panelGlass9.setLayout(new FlowLayout(FlowLayout.LEFT, 5, 9));

        jLabel19.setText("Tgl.Berkas :");
        jLabel19.setPreferredSize(new Dimension(65, 23));
        panelGlass9.add(jLabel19);
        DTPCari1.setForeground(new java.awt.Color(50, 70, 50));
        DTPCari1.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "11-02-2026" }));
        DTPCari1.setDisplayFormat("dd-MM-yyyy");
        DTPCari1.setName("DTPCari1");
        DTPCari1.setOpaque(false);
        DTPCari1.setPreferredSize(new Dimension(95, 23));
        panelGlass9.add(DTPCari1);

        jLabel21.setText("s.d.");
        jLabel21.setPreferredSize(new Dimension(24, 23));
        panelGlass9.add(jLabel21);
        DTPCari2.setForeground(new java.awt.Color(50, 70, 50));
        DTPCari2.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "11-02-2026" }));
        DTPCari2.setDisplayFormat("dd-MM-yyyy");
        DTPCari2.setName("DTPCari2");
        DTPCari2.setOpaque(false);
        DTPCari2.setPreferredSize(new Dimension(95, 23));
        panelGlass9.add(DTPCari2);

        jLabel6.setText("Key Word :");
        jLabel6.setPreferredSize(new Dimension(65, 23));
        panelGlass9.add(jLabel6);

        TCari.setPreferredSize(new Dimension(220, 23));
        TCari.addKeyListener(new java.awt.event.KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent evt) {
                if (evt.getKeyCode() == KeyEvent.VK_ENTER) {
                    BtnCariActionPerformed(null);
                }
            }
        });
        panelGlass9.add(TCari);

        BtnCari.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/accept.png")));
        BtnCari.setMnemonic('1');
        BtnCari.setText("Cari");
        BtnCari.setToolTipText("Alt+1");
        BtnCari.setPreferredSize(new Dimension(70, 23));
        BtnCari.addActionListener(evt -> BtnCariActionPerformed(evt));
        panelGlass9.add(BtnCari);

        jLabel7.setText("Record :");
        jLabel7.setPreferredSize(new Dimension(50, 23));
        panelGlass9.add(jLabel7);

        LCount.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        LCount.setPreferredSize(new Dimension(60, 23));
        panelGlass9.add(LCount);

        jPanel3.add(panelGlass9, BorderLayout.PAGE_START);
        internalFrame1.add(jPanel3, BorderLayout.PAGE_END);

        PanelInput.setPreferredSize(new Dimension(560, 285));
        PanelInput.setLayout(new BorderLayout());

        FormInput.setPreferredSize(new Dimension(560, 285));
        FormInput.setLayout(null);

        jLabel3.setText("No.Surat :");
        jLabel3.setBounds(0, 10, 85, 23);
        FormInput.add(jLabel3);

        NoSurat.setBounds(90, 10, 125, 23);
        FormInput.add(NoSurat);

        jLabel4.setText("No.Rawat :");
        jLabel4.setBounds(220, 10, 65, 23);
        FormInput.add(jLabel4);

        TNoRw.setBounds(290, 10, 110, 23);
        FormInput.add(TNoRw);

        TNoRM.setBounds(405, 10, 70, 23);
        FormInput.add(TNoRM);

        TPasien.setBounds(90, 40, 385, 23);
        FormInput.add(TPasien);

        jLabel16.setText("Tanggal :");
        jLabel16.setBounds(0, 70, 85, 23);
        FormInput.add(jLabel16);

        Tanggal.setForeground(new java.awt.Color(50, 70, 50));
        Tanggal.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "11-02-2026" }));
        Tanggal.setDisplayFormat("dd-MM-yyyy");
        Tanggal.setName("Tanggal");
        Tanggal.setOpaque(false);
        Tanggal.setBounds(90, 70, 110, 23);
        FormInput.add(Tanggal);

        jLabel14.setText("Jam :");
        jLabel14.setBounds(205, 70, 35, 23);
        FormInput.add(jLabel14);

        Jam.setBounds(245, 70, 55, 23);
        FormInput.add(Jam);

        jLabel10.setText("Hubungan :");
        jLabel10.setBounds(305, 70, 70, 23);
        FormInput.add(jLabel10);

        Hubungan.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Diri sendiri", "Ayah", "Ibu", "Suami", "Istri", "Anak", "Lain-lain" }));
        Hubungan.setBounds(380, 70, 95, 23);
        FormInput.add(Hubungan);

        jLabel22.setText("Yang Menyetujui / Pemberi Edukasi");
        jLabel22.setBounds(0, 100, 220, 23);
        FormInput.add(jLabel22);

        jLabel8.setText("Nama :");
        jLabel8.setBounds(0, 130, 85, 23);
        FormInput.add(jLabel8);

        NamaPenyetuju.setBounds(90, 130, 385, 23);
        FormInput.add(NamaPenyetuju);

        jLabel9.setText("Alamat :");
        jLabel9.setBounds(0, 160, 85, 23);
        FormInput.add(jLabel9);

        AlamatPenyetuju.setBounds(90, 160, 385, 23);
        FormInput.add(AlamatPenyetuju);

        jLabel11.setText("Umur :");
        jLabel11.setBounds(0, 190, 85, 23);
        FormInput.add(jLabel11);

        UmurPenyetuju.setBounds(90, 190, 110, 23);
        FormInput.add(UmurPenyetuju);

        jLabel12.setText("J.K. :");
        jLabel12.setBounds(205, 190, 35, 23);
        FormInput.add(jLabel12);

        JKPenyetuju.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Laki-laki", "Perempuan" }));
        JKPenyetuju.setBounds(245, 190, 115, 23);
        FormInput.add(JKPenyetuju);

        jLabel23.setText("Pasien Yang Dirujuk");
        jLabel23.setBounds(0, 220, 150, 23);
        FormInput.add(jLabel23);

        jLabel24.setText("Nama :");
        jLabel24.setBounds(0, 250, 85, 23);
        FormInput.add(jLabel24);

        NamaDirujuk.setBounds(90, 250, 385, 23);
        FormInput.add(NamaDirujuk);

        jLabel25.setText("Alamat :");
        jLabel25.setBounds(0, 280, 85, 23);
        FormInput.add(jLabel25);

        AlamatDirujuk.setBounds(90, 280, 385, 23);
        FormInput.add(AlamatDirujuk);

        jLabel26.setText("Umur :");
        jLabel26.setBounds(0, 310, 85, 23);
        FormInput.add(jLabel26);

        UmurDirujuk.setBounds(90, 310, 110, 23);
        FormInput.add(UmurDirujuk);

        jLabel27.setText("J.K. :");
        jLabel27.setBounds(205, 310, 35, 23);
        FormInput.add(jLabel27);

        JKDirujuk.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Laki-laki", "Perempuan" }));
        JKDirujuk.setBounds(245, 310, 115, 23);
        FormInput.add(JKDirujuk);

        jLabel28.setText("Saksi 1 :");
        jLabel28.setBounds(0, 340, 85, 23);
        FormInput.add(jLabel28);

        Saksi1.setBounds(90, 340, 170, 23);
        FormInput.add(Saksi1);

        jLabel29.setText("Saksi 2 :");
        jLabel29.setBounds(265, 340, 55, 23);
        FormInput.add(jLabel29);

        Saksi2.setBounds(325, 340, 150, 23);
        FormInput.add(Saksi2);

        jLabel17.setText("Petugas :");
        jLabel17.setBounds(0, 370, 85, 23);
        FormInput.add(jLabel17);

        KdPetugas.setBounds(90, 370, 110, 23);
        FormInput.add(KdPetugas);

        NmPetugas.setBounds(205, 370, 235, 23);
        FormInput.add(NmPetugas);

        BtnPetugas.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/190.png")));
        BtnPetugas.setMnemonic('2');
        BtnPetugas.setToolTipText("Alt+2");
        BtnPetugas.setBounds(445, 370, 28, 23);
        BtnPetugas.addActionListener(evt -> BtnPetugasActionPerformed(evt));
        FormInput.add(BtnPetugas);

        ChkInput.setBorder(null);
        ChkInput.setText(".: Input Surat Pernyataan Edukasi Rujukan");
        ChkInput.setPreferredSize(new Dimension(192, 20));
        ChkInput.addActionListener(evt -> ChkInputActionPerformed(evt));
        FormInput.add(ChkInput);
        ChkInput.setBounds(0, 0, 220, 20);

        jLabel5.setText("Pasien :");
        jLabel5.setBounds(0, 40, 85, 23);
        FormInput.add(jLabel5);

        jLabel13.setText("*) Data pasien dirujuk mengikuti pasien yang sedang dipilih.");
        jLabel13.setBounds(0, 400, 300, 23);
        FormInput.add(jLabel13);

        jLabel15.setText("*) Hubungan diisi sesuai pihak yang diberikan edukasi.");
        jLabel15.setBounds(0, 420, 300, 23);
        FormInput.add(jLabel15);

        jLabel18.setText("*) Cetak surat dilakukan dari data yang sudah tersimpan.");
        jLabel18.setBounds(0, 440, 300, 23);
        FormInput.add(jLabel18);

        jLabel20.setText("No.RM");
        jLabel20.setBounds(405, 35, 70, 10);
        FormInput.add(jLabel20);

        jLabel30.setText("Dokter/Petugas");
        jLabel30.setBounds(205, 395, 120, 10);
        FormInput.add(jLabel30);

        PanelInput.add(FormInput, BorderLayout.CENTER);
        internalFrame1.add(PanelInput, BorderLayout.PAGE_START);

        getContentPane().setLayout(new BorderLayout(1, 1));
        getContentPane().add(internalFrame1, BorderLayout.CENTER);
        pack();
    }

    private void BtnSimpanActionPerformed(java.awt.event.ActionEvent evt) {
        if (TNoRw.getText().trim().equals("")) {
            JOptionPane.showMessageDialog(null, "Maaf, silahkan pilih pasien lebih dulu.");
            return;
        }
        if (NamaPenyetuju.getText().trim().equals("")) {
            JOptionPane.showMessageDialog(null, "Nama pemberi pernyataan wajib diisi.");
            NamaPenyetuju.requestFocus();
            return;
        }
        if (AlamatPenyetuju.getText().trim().equals("")) {
            JOptionPane.showMessageDialog(null, "Alamat pemberi pernyataan wajib diisi.");
            AlamatPenyetuju.requestFocus();
            return;
        }
        if (UmurPenyetuju.getText().trim().equals("")) {
            JOptionPane.showMessageDialog(null, "Umur pemberi pernyataan wajib diisi.");
            UmurPenyetuju.requestFocus();
            return;
        }
        if (KdPetugas.getText().trim().equals("")) {
            JOptionPane.showMessageDialog(null, "Petugas pelaksana wajib dipilih.");
            BtnPetugas.requestFocus();
            return;
        }
        if (Jam.getText().trim().equals("")) {
            JOptionPane.showMessageDialog(null, "Jam wajib diisi.");
            Jam.requestFocus();
            return;
        }

        if (Sequel.menyimpantf("surat_pernyataan_edukasi_rujukan", "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?", "Data", 17, new String[]{
            NoSurat.getText(), TNoRw.getText(), TNoRM.getText(), Valid.SetTgl(Tanggal.getSelectedItem() + ""),
            Jam.getText(), NamaPenyetuju.getText(), AlamatPenyetuju.getText(), UmurPenyetuju.getText(),
            JKPenyetuju.getSelectedItem().toString().substring(0, 1), Hubungan.getSelectedItem().toString(),
            NamaDirujuk.getText(), AlamatDirujuk.getText(), UmurDirujuk.getText(),
            JKDirujuk.getSelectedItem().toString().substring(0, 1), Saksi1.getText(), Saksi2.getText(), KdPetugas.getText()
        })) {
            tampil();
            emptTeks();
        }
    }

    private void BtnBatalActionPerformed(java.awt.event.ActionEvent evt) {
        emptTeks();
    }

    private void BtnHapusActionPerformed(java.awt.event.ActionEvent evt) {
        if (tbObat.getSelectedRow() == -1) {
            JOptionPane.showMessageDialog(null, "Pilih data yang ingin dihapus.");
            return;
        }
        if (JOptionPane.showConfirmDialog(null, "Hapus surat yang dipilih?", "Konfirmasi", JOptionPane.YES_NO_OPTION) == JOptionPane.YES_OPTION) {
            hapus();
        }
    }

    private void BtnEditActionPerformed(java.awt.event.ActionEvent evt) {
        if (tbObat.getSelectedRow() == -1) {
            JOptionPane.showMessageDialog(null, "Pilih data yang ingin diganti.");
            return;
        }
        ganti();
    }

    private void BtnPrintActionPerformed(java.awt.event.ActionEvent evt) {
        cetakSurat();
    }

    private void BtnAllActionPerformed(java.awt.event.ActionEvent evt) {
        TCari.setText("");
        tampil();
    }

    private void BtnKeluarActionPerformed(java.awt.event.ActionEvent evt) {
        dispose();
    }

    private void BtnCariActionPerformed(java.awt.event.ActionEvent evt) {
        runBackground(this::tampil);
    }

    private void BtnPetugasActionPerformed(java.awt.event.ActionEvent evt) {
        if (petugas == null || !petugas.isDisplayable()) {
            petugas = new DlgCariPetugas(null, false);
            petugas.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
            petugas.addWindowListener(new WindowAdapter() {
                @Override
                public void windowClosed(WindowEvent e) {
                    if (petugas != null && petugas.getTable().getSelectedRow() != -1) {
                        KdPetugas.setText(petugas.getTable().getValueAt(petugas.getTable().getSelectedRow(), 0).toString());
                        NmPetugas.setText(petugas.getTable().getValueAt(petugas.getTable().getSelectedRow(), 1).toString());
                    }
                    BtnPetugas.requestFocus();
                    petugas = null;
                }
            });
            petugas.setSize(internalFrame1.getWidth() - 20, internalFrame1.getHeight() - 20);
            petugas.setLocationRelativeTo(internalFrame1);
        }
        if (petugas == null) {
            return;
        }
        if (!petugas.isVisible()) {
            petugas.isCek();
            petugas.emptTeks();
            petugas.setVisible(true);
        } else {
            petugas.toFront();
        }
    }

    private void ChkInputActionPerformed(java.awt.event.ActionEvent evt) {
        isForm();
    }

    private void formWindowOpened(java.awt.event.WindowEvent evt) {
        if (koneksiDB.CARICEPAT().equals("aktif")) {
            TCari.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
                @Override
                public void insertUpdate(DocumentEvent e) {
                    if (TCari.getText().length() > 2) {
                        runBackground(SuratPernyataanEdukasiRujukan.this::tampil);
                    }
                }

                @Override
                public void removeUpdate(DocumentEvent e) {
                    if (TCari.getText().length() > 2) {
                        runBackground(SuratPernyataanEdukasiRujukan.this::tampil);
                    }
                }

                @Override
                public void changedUpdate(DocumentEvent e) {
                    if (TCari.getText().length() > 2) {
                        runBackground(SuratPernyataanEdukasiRujukan.this::tampil);
                    }
                }
            });
        }
    }

    private void tampil() {
        Valid.tabelKosong(tabMode);
        try {
            if (TCari.getText().trim().equals("")) {
                ps = koneksi.prepareStatement(
                    "select surat_pernyataan_edukasi_rujukan.no_surat,surat_pernyataan_edukasi_rujukan.no_rawat," +
                    "surat_pernyataan_edukasi_rujukan.no_rkm_medis,surat_pernyataan_edukasi_rujukan.nama_pasien_rujukan," +
                    "surat_pernyataan_edukasi_rujukan.umur_pasien_rujukan,surat_pernyataan_edukasi_rujukan.jk_pasien_rujukan," +
                    "surat_pernyataan_edukasi_rujukan.tanggal,surat_pernyataan_edukasi_rujukan.jam," +
                    "surat_pernyataan_edukasi_rujukan.nama_penyetuju,surat_pernyataan_edukasi_rujukan.alamat_penyetuju," +
                    "surat_pernyataan_edukasi_rujukan.umur_penyetuju,surat_pernyataan_edukasi_rujukan.jk_penyetuju," +
                    "surat_pernyataan_edukasi_rujukan.hubungan,surat_pernyataan_edukasi_rujukan.nama_pasien_rujukan," +
                    "surat_pernyataan_edukasi_rujukan.alamat_pasien_rujukan,surat_pernyataan_edukasi_rujukan.umur_pasien_rujukan," +
                    "surat_pernyataan_edukasi_rujukan.jk_pasien_rujukan,surat_pernyataan_edukasi_rujukan.saksi1," +
                    "surat_pernyataan_edukasi_rujukan.saksi2,surat_pernyataan_edukasi_rujukan.nip,petugas.nama " +
                    "from surat_pernyataan_edukasi_rujukan inner join petugas on surat_pernyataan_edukasi_rujukan.nip=petugas.nip " +
                    "where surat_pernyataan_edukasi_rujukan.tanggal between ? and ? order by surat_pernyataan_edukasi_rujukan.tanggal,surat_pernyataan_edukasi_rujukan.jam");
            } else {
                ps = koneksi.prepareStatement(
                    "select surat_pernyataan_edukasi_rujukan.no_surat,surat_pernyataan_edukasi_rujukan.no_rawat," +
                    "surat_pernyataan_edukasi_rujukan.no_rkm_medis,surat_pernyataan_edukasi_rujukan.nama_pasien_rujukan," +
                    "surat_pernyataan_edukasi_rujukan.umur_pasien_rujukan,surat_pernyataan_edukasi_rujukan.jk_pasien_rujukan," +
                    "surat_pernyataan_edukasi_rujukan.tanggal,surat_pernyataan_edukasi_rujukan.jam," +
                    "surat_pernyataan_edukasi_rujukan.nama_penyetuju,surat_pernyataan_edukasi_rujukan.alamat_penyetuju," +
                    "surat_pernyataan_edukasi_rujukan.umur_penyetuju,surat_pernyataan_edukasi_rujukan.jk_penyetuju," +
                    "surat_pernyataan_edukasi_rujukan.hubungan,surat_pernyataan_edukasi_rujukan.nama_pasien_rujukan," +
                    "surat_pernyataan_edukasi_rujukan.alamat_pasien_rujukan,surat_pernyataan_edukasi_rujukan.umur_pasien_rujukan," +
                    "surat_pernyataan_edukasi_rujukan.jk_pasien_rujukan,surat_pernyataan_edukasi_rujukan.saksi1," +
                    "surat_pernyataan_edukasi_rujukan.saksi2,surat_pernyataan_edukasi_rujukan.nip,petugas.nama " +
                    "from surat_pernyataan_edukasi_rujukan inner join petugas on surat_pernyataan_edukasi_rujukan.nip=petugas.nip " +
                    "where surat_pernyataan_edukasi_rujukan.tanggal between ? and ? and " +
                    "(surat_pernyataan_edukasi_rujukan.no_surat like ? or surat_pernyataan_edukasi_rujukan.no_rawat like ? or " +
                    "surat_pernyataan_edukasi_rujukan.no_rkm_medis like ? or surat_pernyataan_edukasi_rujukan.nama_penyetuju like ? or " +
                    "surat_pernyataan_edukasi_rujukan.nama_pasien_rujukan like ? or petugas.nama like ?) " +
                    "order by surat_pernyataan_edukasi_rujukan.tanggal,surat_pernyataan_edukasi_rujukan.jam");
            }

            try {
                ps.setString(1, Valid.SetTgl(DTPCari1.getSelectedItem() + ""));
                ps.setString(2, Valid.SetTgl(DTPCari2.getSelectedItem() + ""));
                if (!TCari.getText().trim().equals("")) {
                    ps.setString(3, "%" + TCari.getText().trim() + "%");
                    ps.setString(4, "%" + TCari.getText().trim() + "%");
                    ps.setString(5, "%" + TCari.getText().trim() + "%");
                    ps.setString(6, "%" + TCari.getText().trim() + "%");
                    ps.setString(7, "%" + TCari.getText().trim() + "%");
                    ps.setString(8, "%" + TCari.getText().trim() + "%");
                }
                rs = ps.executeQuery();
                while (rs.next()) {
                    tabMode.addRow(new Object[]{
                        rs.getString(1), rs.getString(2), rs.getString(3), rs.getString(4), rs.getString(5), rs.getString(6),
                        rs.getString(7), rs.getString(8), rs.getString(9), rs.getString(10), rs.getString(11), rs.getString(12),
                        rs.getString(13), rs.getString(14), rs.getString(15), rs.getString(16), rs.getString(17), rs.getString(18),
                        rs.getString(19), rs.getString(20), rs.getString(21)
                    });
                }
            } catch (Exception e) {
                System.out.println("Notif : " + e);
            } finally {
                if (rs != null) {
                    rs.close();
                }
                if (ps != null) {
                    ps.close();
                }
            }
        } catch (Exception e) {
            System.out.println("Notifikasi : " + e);
        }
        LCount.setText(String.valueOf(tabMode.getRowCount()));
    }

    public void emptTeks() {
        NamaPenyetuju.setText("");
        AlamatPenyetuju.setText("");
        UmurPenyetuju.setText("");
        JKPenyetuju.setSelectedIndex(0);
        Hubungan.setSelectedIndex(0);
        Saksi1.setText("");
        Saksi2.setText("");
        Tanggal.setDate(new Date());
        Jam.setText(jamFormat.format(new Date()));
        Valid.autoNomer3(
            "select ifnull(MAX(CONVERT(RIGHT(surat_pernyataan_edukasi_rujukan.no_surat,3),signed)),0) from surat_pernyataan_edukasi_rujukan " +
            "where surat_pernyataan_edukasi_rujukan.tanggal='" + Valid.SetTgl(Tanggal.getSelectedItem() + "") + "' ",
            "SER" + Tanggal.getSelectedItem().toString().substring(6, 10) + Tanggal.getSelectedItem().toString().substring(3, 5) + Tanggal.getSelectedItem().toString().substring(0, 2),
            3, NoSurat
        );
        isRawat();
        NamaPenyetuju.requestFocus();
    }

    private void getData() {
        if (tbObat.getSelectedRow() != -1) {
            NoSurat.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 0).toString());
            TNoRw.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 1).toString());
            TNoRM.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 2).toString());
            TPasien.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 3).toString());
            Umur.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 4).toString());
            JK.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 5).toString());
            Valid.SetTgl(Tanggal, tbObat.getValueAt(tbObat.getSelectedRow(), 6).toString());
            Jam.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 7).toString());
            NamaPenyetuju.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 8).toString());
            AlamatPenyetuju.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 9).toString());
            UmurPenyetuju.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 10).toString());
            JKPenyetuju.setSelectedItem(tbObat.getValueAt(tbObat.getSelectedRow(), 11).toString().equals("P") ? "Perempuan" : "Laki-laki");
            Hubungan.setSelectedItem(tbObat.getValueAt(tbObat.getSelectedRow(), 12).toString());
            NamaDirujuk.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 13).toString());
            AlamatDirujuk.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 14).toString());
            UmurDirujuk.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 15).toString());
            JKDirujuk.setSelectedItem(tbObat.getValueAt(tbObat.getSelectedRow(), 16).toString().equals("P") ? "Perempuan" : "Laki-laki");
            Saksi1.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 17).toString());
            Saksi2.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 18).toString());
            KdPetugas.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 19).toString());
            NmPetugas.setText(tbObat.getValueAt(tbObat.getSelectedRow(), 20).toString());
        }
    }

    private void isRawat() {
        if (TNoRw.getText().trim().equals("")) {
            return;
        }
        try {
            ps = koneksi.prepareStatement(
                "select reg_periksa.no_rkm_medis,pasien.nm_pasien,pasien.jk,concat(reg_periksa.umurdaftar,' ',reg_periksa.sttsumur) as umur," +
                "concat(pasien.alamat,', ',kelurahan.nm_kel,', ',kecamatan.nm_kec,', ',kabupaten.nm_kab,', ',propinsi.nm_prop) as alamat " +
                "from reg_periksa inner join pasien on reg_periksa.no_rkm_medis=pasien.no_rkm_medis " +
                "inner join kelurahan on pasien.kd_kel=kelurahan.kd_kel " +
                "inner join kecamatan on pasien.kd_kec=kecamatan.kd_kec " +
                "inner join kabupaten on pasien.kd_kab=kabupaten.kd_kab " +
                "inner join propinsi on pasien.kd_prop=propinsi.kd_prop where reg_periksa.no_rawat=?");
            try {
                ps.setString(1, TNoRw.getText());
                rs = ps.executeQuery();
                if (rs.next()) {
                    TNoRM.setText(rs.getString("no_rkm_medis"));
                    TPasien.setText(rs.getString("nm_pasien"));
                    Umur.setText(rs.getString("umur"));
                    JK.setText(rs.getString("jk"));
                    NamaDirujuk.setText(rs.getString("nm_pasien"));
                    AlamatDirujuk.setText(rs.getString("alamat"));
                    UmurDirujuk.setText(rs.getString("umur"));
                    JKDirujuk.setSelectedItem(rs.getString("jk").equals("P") ? "Perempuan" : "Laki-laki");
                    if (NamaPenyetuju.getText().trim().equals("")) {
                        NamaPenyetuju.setText(rs.getString("nm_pasien"));
                    }
                    if (AlamatPenyetuju.getText().trim().equals("")) {
                        AlamatPenyetuju.setText(rs.getString("alamat"));
                    }
                    if (UmurPenyetuju.getText().trim().equals("")) {
                        UmurPenyetuju.setText(rs.getString("umur"));
                    }
                    JKPenyetuju.setSelectedItem(rs.getString("jk").equals("P") ? "Perempuan" : "Laki-laki");
                }
            } catch (Exception e) {
                System.out.println("Notif : " + e);
            } finally {
                if (rs != null) {
                    rs.close();
                }
                if (ps != null) {
                    ps.close();
                }
            }
        } catch (Exception e) {
            System.out.println("Notif : " + e);
        }
    }

    public void setNoRm(String norwt, Date tgl2) {
        TNoRw.setText(norwt);
        TCari.setText(norwt);
        DTPCari2.setDate(tgl2);
        isRawat();
        ChkInput.setSelected(true);
        isForm();
        runBackground(this::tampil);
    }

    private void isForm() {
        PanelInput.setPreferredSize(new Dimension(560, 470));
        FormInput.setVisible(true);
    }

    public void isCek() {
        BtnSimpan.setEnabled(akses.getsurat_pernyataan_pasien_umum());
        BtnHapus.setEnabled(akses.getsurat_pernyataan_pasien_umum());
        BtnEdit.setEnabled(akses.getsurat_pernyataan_pasien_umum());
        BtnPrint.setEnabled(akses.getsurat_pernyataan_pasien_umum());
        if (akses.getjml2() >= 1) {
            KdPetugas.setEditable(false);
            BtnPetugas.setEnabled(false);
            KdPetugas.setText(akses.getkode());
            NmPetugas.setText(Sequel.CariPetugas(KdPetugas.getText()));
            if (NmPetugas.getText().equals("")) {
                KdPetugas.setText("");
                JOptionPane.showMessageDialog(null, "User login bukan petugas...!!");
            }
        }
    }

    private void ganti() {
        if (Sequel.mengedittf("surat_pernyataan_edukasi_rujukan", "no_surat=?", "no_surat=?,no_rawat=?,no_rkm_medis=?,tanggal=?,jam=?,nama_penyetuju=?,alamat_penyetuju=?,umur_penyetuju=?,jk_penyetuju=?,hubungan=?,nama_pasien_rujukan=?,alamat_pasien_rujukan=?,umur_pasien_rujukan=?,jk_pasien_rujukan=?,saksi1=?,saksi2=?,nip=?", 18, new String[]{
            NoSurat.getText(), TNoRw.getText(), TNoRM.getText(), Valid.SetTgl(Tanggal.getSelectedItem() + ""), Jam.getText(),
            NamaPenyetuju.getText(), AlamatPenyetuju.getText(), UmurPenyetuju.getText(),
            JKPenyetuju.getSelectedItem().toString().substring(0, 1), Hubungan.getSelectedItem().toString(),
            NamaDirujuk.getText(), AlamatDirujuk.getText(), UmurDirujuk.getText(),
            JKDirujuk.getSelectedItem().toString().substring(0, 1), Saksi1.getText(), Saksi2.getText(), KdPetugas.getText(),
            tbObat.getValueAt(tbObat.getSelectedRow(), 0).toString()
        })) {
            tampil();
            emptTeks();
        }
    }

    private void hapus() {
        if (Sequel.queryu2tf("delete from surat_pernyataan_edukasi_rujukan where no_surat=?", 1, new String[]{
            tbObat.getValueAt(tbObat.getSelectedRow(), 0).toString()
        })) {
            tabMode.removeRow(tbObat.getSelectedRow());
            LCount.setText(String.valueOf(tabMode.getRowCount()));
            emptTeks();
        } else {
            JOptionPane.showMessageDialog(null, "Gagal menghapus..!!");
        }
    }

    private void cetakSurat() {
        if (tbObat.getSelectedRow() == -1) {
            JOptionPane.showMessageDialog(null, "Pilih surat yang akan dicetak.");
            return;
        }
        try {
            siapkanReport();
            Map<String, Object> param = new HashMap<>();
            param.put("namars", akses.getnamars());
            param.put("alamatrs", akses.getalamatrs());
            param.put("kotars", akses.getkabupatenrs());
            param.put("propinsirs", akses.getpropinsirs());
            param.put("kontakrs", akses.getkontakrs());
            param.put("emailrs", akses.getemailrs());
            param.put("logo", Sequel.cariGambar("select setting.logo from setting"));
            Valid.MyReportqry(
                "rptSuratPernyataanEdukasiRujukan.jasper",
                "report",
                "::[ Surat Pernyataan Edukasi Rujukan ]::",
                "select surat_pernyataan_edukasi_rujukan.no_surat,surat_pernyataan_edukasi_rujukan.no_rawat," +
                "surat_pernyataan_edukasi_rujukan.no_rkm_medis,surat_pernyataan_edukasi_rujukan.tanggal," +
                "surat_pernyataan_edukasi_rujukan.jam,surat_pernyataan_edukasi_rujukan.nama_penyetuju," +
                "surat_pernyataan_edukasi_rujukan.alamat_penyetuju,surat_pernyataan_edukasi_rujukan.umur_penyetuju," +
                "surat_pernyataan_edukasi_rujukan.jk_penyetuju,surat_pernyataan_edukasi_rujukan.hubungan," +
                "surat_pernyataan_edukasi_rujukan.nama_pasien_rujukan,surat_pernyataan_edukasi_rujukan.alamat_pasien_rujukan," +
                "surat_pernyataan_edukasi_rujukan.umur_pasien_rujukan,surat_pernyataan_edukasi_rujukan.jk_pasien_rujukan," +
                "surat_pernyataan_edukasi_rujukan.saksi1,surat_pernyataan_edukasi_rujukan.saksi2," +
                "surat_pernyataan_edukasi_rujukan.nip,petugas.nama from surat_pernyataan_edukasi_rujukan " +
                "inner join petugas on surat_pernyataan_edukasi_rujukan.nip=petugas.nip " +
                "where surat_pernyataan_edukasi_rujukan.no_surat='" + tbObat.getValueAt(tbObat.getSelectedRow(), 0).toString() + "'",
                param
            );
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "Gagal menyiapkan report surat: " + e.getMessage());
            System.out.println("Notif : " + e);
        }
    }

    private void siapkanReport() throws Exception {
        File jrxml = new File("./report/rptSuratPernyataanEdukasiRujukan.jrxml");
        File jasper = new File("./report/rptSuratPernyataanEdukasiRujukan.jasper");
        if (!jrxml.exists()) {
            throw new Exception("File JRXML report tidak ditemukan.");
        }
        if (!jasper.exists() || jrxml.lastModified() > jasper.lastModified()) {
            String compilerClass = JRJavacCompiler.class.getName();
            SimpleJasperReportsContext jasperContext =
                new SimpleJasperReportsContext(DefaultJasperReportsContext.getInstance());
            jasperContext.setProperty(JRCompiler.COMPILER_CLASS, compilerClass);
            jasperContext.setProperty(JRCompiler.COMPILER_PREFIX + "java", compilerClass);
            String classpath = System.getProperty("java.class.path");
            if (classpath != null && !classpath.trim().isEmpty()) {
                jasperContext.setProperty(JRCompiler.COMPILER_CLASSPATH, classpath);
            }
            JasperCompileManager.getInstance(jasperContext).compileToFile(jrxml.getPath(), jasper.getPath());
        }
    }

    private void cekTabel() {
        try (Statement st = koneksi.createStatement()) {
            st.executeUpdate(
                "CREATE TABLE IF NOT EXISTS surat_pernyataan_edukasi_rujukan (" +
                "no_surat varchar(20) NOT NULL," +
                "no_rawat varchar(17) NOT NULL," +
                "no_rkm_medis varchar(15) NOT NULL," +
                "tanggal date NOT NULL," +
                "jam varchar(5) NOT NULL," +
                "nama_penyetuju varchar(50) NOT NULL," +
                "alamat_penyetuju varchar(150) NOT NULL," +
                "umur_penyetuju varchar(20) NOT NULL," +
                "jk_penyetuju enum('L','P') NOT NULL," +
                "hubungan varchar(30) NOT NULL," +
                "nama_pasien_rujukan varchar(50) NOT NULL," +
                "alamat_pasien_rujukan varchar(150) NOT NULL," +
                "umur_pasien_rujukan varchar(20) NOT NULL," +
                "jk_pasien_rujukan enum('L','P') NOT NULL," +
                "saksi1 varchar(50) NOT NULL," +
                "saksi2 varchar(50) NOT NULL," +
                "nip varchar(20) NOT NULL," +
                "PRIMARY KEY (no_surat)," +
                "KEY no_rawat (no_rawat)," +
                "KEY nip (nip)," +
                "CONSTRAINT surat_pernyataan_edukasi_rujukan_ibfk_1 FOREIGN KEY (no_rawat) REFERENCES reg_periksa (no_rawat) ON DELETE CASCADE ON UPDATE CASCADE," +
                "CONSTRAINT surat_pernyataan_edukasi_rujukan_ibfk_2 FOREIGN KEY (nip) REFERENCES petugas (nip) ON DELETE CASCADE ON UPDATE CASCADE" +
                ") ENGINE=InnoDB DEFAULT CHARSET=latin1"
            );
        } catch (Exception e) {
            System.out.println("Notif cek tabel surat_pernyataan_edukasi_rujukan : " + e);
        }
    }

    private void runBackground(Runnable task) {
        if (ceksukses) {
            return;
        }
        if (executor.isShutdown() || executor.isTerminated()) {
            return;
        }
        if (!isDisplayable()) {
            return;
        }

        ceksukses = true;
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        try {
            executor.submit(() -> {
                try {
                    task.run();
                } finally {
                    ceksukses = false;
                    SwingUtilities.invokeLater(() -> {
                        if (isDisplayable()) {
                            setCursor(Cursor.getDefaultCursor());
                        }
                    });
                }
            });
        } catch (RejectedExecutionException ex) {
            ceksukses = false;
        }
    }

    @Override
    public void dispose() {
        executor.shutdownNow();
        super.dispose();
    }

    public static void main(String args[]) {
        java.awt.EventQueue.invokeLater(() -> {
            SuratPernyataanEdukasiRujukan dialog = new SuratPernyataanEdukasiRujukan(new javax.swing.JFrame(), true);
            dialog.addWindowListener(new java.awt.event.WindowAdapter() {
                @Override
                public void windowClosing(java.awt.event.WindowEvent e) {
                    System.exit(0);
                }
            });
            dialog.setVisible(true);
        });
    }
}
