package laporan;

import fungsi.akses;
import fungsi.koneksiDB;
import fungsi.validasi;
import java.awt.Color;
import java.awt.BorderLayout;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.KeyEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.SwingWorker;
import javax.swing.SwingUtilities;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.Document;
import javax.swing.text.html.HTMLEditorKit;
import javax.swing.text.html.StyleSheet;
import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class DlgLaporanKegiatanPelayanan extends javax.swing.JDialog {
    private final validasi Valid = new validasi();
    private final Connection koneksi = koneksiDB.condb();
    private final List<BarisLaporan> dataLaporan = new ArrayList<BarisLaporan>();
    private StringBuilder htmlContent;
    private PreparedStatement ps;
    private ResultSet rs;
    private double totalSkbsBpjs = 0;
    private double totalSkbsUmum = 0;
    private double totalTindakanBpjs = 0;
    private double totalTindakanUmum = 0;
    private double totalHarga = 0;
    private double totalUsgBpjs = 0;
    private double totalUsgUmum = 0;
    private double totalObatBpjs = 0;
    private double totalObatGigiUmum = 0;
    private double totalObatPoliUmum = 0;
    private int totalPasienBpjs = 0;
    private int totalPasienUmum = 0;
    private boolean sedangMemuat = false;

    public DlgLaporanKegiatanPelayanan(java.awt.Frame parent, boolean modal) {
        super(parent, modal);
        initComponents();
        this.setLocation(8, 1);
        setSize(1200, 700);
    }

    @SuppressWarnings("unchecked")
    private void initComponents() {
        internalFrame1 = new widget.InternalFrame();
        panelFilter = new widget.panelisi();
        panelFilterInput = new widget.panelisi();
        panelAksi = new widget.panelisi();
        label11 = new widget.Label();
        Tgl1 = new widget.Tanggal();
        label18 = new widget.Label();
        Tgl2 = new widget.Tanggal();
        labelDokter = new widget.Label();
        cmbDokter = new widget.ComboBox();
        labelUnit = new widget.Label();
        cmbUnit = new widget.ComboBox();
        labelShift = new widget.Label();
        cmbShift = new widget.ComboBox();
        labelKeyword = new widget.Label();
        TCari = new widget.TextBox();
        btnTampilkan = new widget.Button();
        btnPeriodeBulan = new widget.Button();
        labelInfo = new widget.Label();
        BtnExport = new widget.Button();
        BtnPrint = new widget.Button();
        BtnKeluar = new widget.Button();
        Scroll = new widget.ScrollPane();
        LoadHTML = new widget.editorpane();

        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setUndecorated(true);
        setResizable(false);
        addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowOpened(java.awt.event.WindowEvent evt) {
                formWindowOpened(evt);
            }
        });

        internalFrame1.setBorder(javax.swing.BorderFactory.createTitledBorder(
                javax.swing.BorderFactory.createLineBorder(new java.awt.Color(240, 245, 235)),
                "::[ Laporan Kegiatan Selama Pelayanan ]::",
                javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION,
                javax.swing.border.TitledBorder.DEFAULT_POSITION,
                new java.awt.Font("Tahoma", 0, 11),
                new java.awt.Color(50, 50, 50)
        ));
        internalFrame1.setName("internalFrame1");
        internalFrame1.setLayout(new BorderLayout(1, 1));

        panelFilter.setName("panelFilter");
        panelFilter.setPreferredSize(new Dimension(100, 92));
        panelFilter.setLayout(new BorderLayout(0, 6));
        panelFilter.setBorder(new EmptyBorder(8, 10, 8, 10));
        panelFilter.setBackground(Color.WHITE);

        panelFilterInput.setName("panelFilterInput");
        panelFilterInput.setLayout(new FlowLayout(FlowLayout.LEFT, 5, 0));
        panelFilterInput.setBackground(Color.WHITE);

        label11.setText("Periode :");
        label11.setPreferredSize(new Dimension(55, 23));
        panelFilterInput.add(label11);

        Tgl1.setDisplayFormat("dd-MM-yyyy");
        Tgl1.setPreferredSize(new Dimension(95, 23));
        panelFilterInput.add(Tgl1);

        label18.setHorizontalAlignment(SwingConstants.CENTER);
        label18.setText("s.d.");
        label18.setPreferredSize(new Dimension(25, 23));
        panelFilterInput.add(label18);

        Tgl2.setDisplayFormat("dd-MM-yyyy");
        Tgl2.setPreferredSize(new Dimension(95, 23));
        panelFilterInput.add(Tgl2);

        labelDokter.setText("Dokter :");
        labelDokter.setPreferredSize(new Dimension(50, 23));
        panelFilterInput.add(labelDokter);

        cmbDokter.setModel(new DefaultComboBoxModel(new String[] {
            "Semua Dokter", "dr. Placeholder Poli Umum", "dr. Placeholder Poli Gigi"
        }));
        cmbDokter.setPreferredSize(new Dimension(210, 23));
        panelFilterInput.add(cmbDokter);

        labelUnit.setText("Unit :");
        labelUnit.setPreferredSize(new Dimension(35, 23));
        panelFilterInput.add(labelUnit);

        cmbUnit.setModel(new DefaultComboBoxModel(new String[] {
            "Semua Unit", "Poliklinik Umum", "Poliklinik Gigi"
        }));
        cmbUnit.setPreferredSize(new Dimension(160, 23));
        panelFilterInput.add(cmbUnit);

        labelShift.setText("Shift :");
        labelShift.setPreferredSize(new Dimension(40, 23));
        panelFilterInput.add(labelShift);

        cmbShift.setModel(new DefaultComboBoxModel(new String[] {
            "Semua Shift", "Shift 1 (08:00 - 15:00)", "Shift 2 (15:00 - 22:00)"
        }));
        cmbShift.setPreferredSize(new Dimension(165, 23));
        panelFilterInput.add(cmbShift);

        labelKeyword.setText("Keyword :");
        labelKeyword.setPreferredSize(new Dimension(55, 23));
        panelFilterInput.add(labelKeyword);

        TCari.setPreferredSize(new Dimension(150, 23));
        TCari.addKeyListener(new java.awt.event.KeyAdapter() {
            @Override
            public void keyPressed(java.awt.event.KeyEvent evt) {
                TCariKeyPressed(evt);
            }
        });
        panelFilterInput.add(TCari);

        btnTampilkan.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/accept.png")));
        btnTampilkan.setMnemonic('2');
        btnTampilkan.setToolTipText("Alt+2");
        btnTampilkan.setPreferredSize(new Dimension(28, 23));
        btnTampilkan.addActionListener(this::btnTampilkanActionPerformed);
        panelFilterInput.add(btnTampilkan);

        btnPeriodeBulan.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Search-16x16.png")));
        btnPeriodeBulan.setText("Bulan Ini");
        btnPeriodeBulan.setPreferredSize(new Dimension(100, 23));
        btnPeriodeBulan.addActionListener(this::btnPeriodeBulanActionPerformed);
        panelFilterInput.add(btnPeriodeBulan);

        panelAksi.setName("panelAksi");
        panelAksi.setLayout(new GridBagLayout());
        panelAksi.setBackground(new Color(250, 252, 247));
        panelAksi.setBorder(javax.swing.BorderFactory.createCompoundBorder(
            javax.swing.BorderFactory.createLineBorder(new Color(232, 238, 228)),
            new EmptyBorder(6, 8, 6, 8)
        ));

        labelInfo.setHorizontalAlignment(SwingConstants.LEFT);
        labelInfo.setText("Preview UI siap direview");
        labelInfo.setPreferredSize(new Dimension(220, 23));

        BtnExport.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Search-16x16.png")));
        BtnExport.setText("Export");
        BtnExport.setPreferredSize(new Dimension(100, 28));
        BtnExport.addActionListener(this::BtnExportActionPerformed);

        BtnPrint.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/b_print.png")));
        BtnPrint.setText("Cetak");
        BtnPrint.setPreferredSize(new Dimension(100, 28));
        BtnPrint.addActionListener(this::BtnPrintActionPerformed);

        BtnKeluar.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/exit.png")));
        BtnKeluar.setText("Keluar");
        BtnKeluar.setPreferredSize(new Dimension(100, 28));
        BtnKeluar.addActionListener(this::BtnKeluarActionPerformed);

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 1.0;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.anchor = GridBagConstraints.WEST;
        gbc.insets = new Insets(0, 0, 0, 12);
        panelAksi.add(labelInfo, gbc);

        gbc = new GridBagConstraints();
        gbc.gridx = 1;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.EAST;
        gbc.insets = new Insets(0, 0, 0, 6);
        panelAksi.add(BtnExport, gbc);

        gbc = new GridBagConstraints();
        gbc.gridx = 2;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.EAST;
        gbc.insets = new Insets(0, 0, 0, 6);
        panelAksi.add(BtnPrint, gbc);

        gbc = new GridBagConstraints();
        gbc.gridx = 3;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.EAST;
        panelAksi.add(BtnKeluar, gbc);

        panelFilter.add(panelFilterInput, BorderLayout.NORTH);
        panelFilter.add(panelAksi, BorderLayout.CENTER);

        internalFrame1.add(panelFilter, BorderLayout.PAGE_END);

        Scroll.setName("Scroll");
        Scroll.setOpaque(true);

        LoadHTML.setBorder(null);
        LoadHTML.setName("LoadHTML");
        Scroll.setViewportView(LoadHTML);

        internalFrame1.add(Scroll, BorderLayout.CENTER);

        getContentPane().add(internalFrame1, BorderLayout.CENTER);

        pack();
    }

    private void formWindowOpened(java.awt.event.WindowEvent evt) {
        HTMLEditorKit kit = new HTMLEditorKit();
        LoadHTML.setEditable(false);
        LoadHTML.setEditorKit(kit);
        StyleSheet styleSheet = kit.getStyleSheet();
        styleSheet.addRule(
            "table{border-collapse:collapse;font-family:tahoma;font-size:11px;color:#323232;}"+
            ".isi td{border:1px solid #dfe5d8;height:23px;padding:4px;background:#ffffff;}"+
            ".head td{border:1px solid #dfe5d8;height:26px;padding:4px;background:#f3f7ef;font-weight:bold;text-align:center;}"+
            ".head2 td{border:1px solid #dfe5d8;height:24px;padding:4px;background:#fbfcf8;font-weight:bold;text-align:center;}"+
            ".judul td{padding:3px;border:none;text-align:center;}"+
            ".subjudul td{padding:2px 4px;border:none;text-align:left;color:#4f5f4d;font-size:10px;}"+
            ".kosong{color:#9aa29b;}"+
            ".kanan{text-align:right;}"+
            ".tengah{text-align:center;}"
        );
        Document doc = kit.createDefaultDocument();
        LoadHTML.setDocument(doc);
        aturPeriodeHari(LocalDate.now());
        loadDokter();
        loadUnit();
        tampilkanTemplateAsync();
    }

    private void btnTampilkanActionPerformed(java.awt.event.ActionEvent evt) {
        tampilkanTemplateAsync();
    }

    private void btnPeriodeBulanActionPerformed(java.awt.event.ActionEvent evt) {
        aturPeriodeBulan(LocalDate.now());
        tampilkanTemplateAsync();
    }

    private void BtnKeluarActionPerformed(java.awt.event.ActionEvent evt) {
        dispose();
    }

    private void BtnPrintActionPerformed(java.awt.event.ActionEvent evt) {
        this.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        try {
            File css = new File("file2.css");
            BufferedWriter bg = new BufferedWriter(new FileWriter(css));
            bg.write(
                "table{border-collapse:collapse;font-family:tahoma;font-size:11px;color:#323232;}"+
                ".isi td{border:1px solid #dfe5d8;height:23px;padding:4px;background:#ffffff;}"+
                ".head td{border:1px solid #dfe5d8;height:26px;padding:4px;background:#f3f7ef;font-weight:bold;text-align:center;}"+
                ".head2 td{border:1px solid #dfe5d8;height:24px;padding:4px;background:#fbfcf8;font-weight:bold;text-align:center;}"+
                ".judul td{padding:3px;border:none;text-align:center;}"+
                ".subjudul td{padding:2px 4px;border:none;text-align:left;color:#4f5f4d;font-size:10px;}"+
                ".kosong{color:#9aa29b;}"+
                ".kanan{text-align:right;}"+
                ".tengah{text-align:center;}"
            );
            bg.close();

            File html = new File("LaporanKegiatanPelayananUI.html");
            BufferedWriter bw = new BufferedWriter(new FileWriter(html));
            bw.write(LoadHTML.getText().replaceAll(
                "<head>",
                "<head><link href=\"file2.css\" rel=\"stylesheet\" type=\"text/css\" />"
            ));
            bw.close();
            Desktop.getDesktop().browse(html.toURI());
        } catch (Exception e) {
            System.out.println("Notifikasi : " + e);
        }
        this.setCursor(Cursor.getDefaultCursor());
    }

    private void BtnExportActionPerformed(java.awt.event.ActionEvent evt) {
        exportKeExcel();
    }

    private void TCariKeyPressed(java.awt.event.KeyEvent evt) {
        if (evt.getKeyCode() == KeyEvent.VK_ENTER) {
            tampilkanTemplateAsync();
        } else if (evt.getKeyCode() == KeyEvent.VK_PAGE_DOWN) {
            btnTampilkan.requestFocus();
        } else if (evt.getKeyCode() == KeyEvent.VK_PAGE_UP) {
            BtnKeluar.requestFocus();
        }
    }

    private void aturPeriodeHari(LocalDate acuan) {
        Tgl1.setDate(Date.valueOf(acuan));
        Tgl2.setDate(Date.valueOf(acuan));
    }

    private void aturPeriodeBulan(LocalDate acuan) {
        LocalDate awal = acuan.withDayOfMonth(1);
        LocalDate akhir = acuan.withDayOfMonth(acuan.lengthOfMonth());
        Tgl1.setDate(Date.valueOf(awal));
        Tgl2.setDate(Date.valueOf(akhir));
    }

    private void tampilkanTemplateAsync() {
        if (sedangMemuat) {
            return;
        }
        final FilterLaporan filter = ambilFilterSaatIni();
        setLoadingState(true, "Memuat laporan kegiatan pelayanan...");
        new SwingWorker<ReportSnapshot, Void>() {
            @Override
            protected ReportSnapshot doInBackground() {
                return buatSnapshotLaporan(filter);
            }

            @Override
            protected void done() {
                try {
                    terapkanSnapshot(get());
                    labelInfo.setText("Data tampil: " + dataLaporan.size() + " baris");
                } catch (Exception e) {
                    labelInfo.setText("Gagal memuat data");
                    LoadHTML.setText("<html><body><center>Gagal memuat data : " + e.getMessage() + "</center></body></html>");
                    System.out.println("Notifikasi : " + e);
                } finally {
                    setLoadingState(false, labelInfo.getText());
                }
            }
        }.execute();
    }

    private ReportSnapshot buatSnapshotLaporan(FilterLaporan filter) {
        StringBuilder htmlBuilder = new StringBuilder();
        DateTimeFormatter inputFormat = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        DateTimeFormatter tampilFormat = DateTimeFormatter.ofPattern("dd-MM-yyyy");
        LocalDate awal = LocalDate.parse(filter.tanggalAwal, inputFormat);
        LocalDate akhir = LocalDate.parse(filter.tanggalAkhir, inputFormat);
        if (akhir.isBefore(awal)) {
            LocalDate temp = awal;
            awal = akhir;
            akhir = temp;
        }
        String tanggalAwalSql = inputFormat.format(awal);
        String tanggalAkhirSql = inputFormat.format(akhir);

        htmlBuilder.append("<html><head></head><body>");
        htmlBuilder.append("<table width='2300px' align='center' cellpadding='0' cellspacing='0'>");
        htmlBuilder.append("<tr class='judul'><td colspan='17'>");
        htmlBuilder.append("<font size='4'>").append(akses.getnamars()).append("</font><br>");
        htmlBuilder.append(akses.getalamatrs()).append(", ").append(akses.getkabupatenrs()).append(", ").append(akses.getpropinsirs()).append("<br>");
        htmlBuilder.append(akses.getkontakrs()).append(", E-mail : ").append(akses.getemailrs()).append("<br><br>");
        htmlBuilder.append("<font size='3'>LAPORAN KEGIATAN SELAMA PELAYANAN</font><br>");
        htmlBuilder.append("<font size='2'>Periode ").append(tampilFormat.format(awal)).append(" s.d. ").append(tampilFormat.format(akhir)).append("</font><br>");
        htmlBuilder.append("<font size='2'>Dokter : ").append(escapeHtml(filter.namaDokter))
                .append(" | Unit : ").append(escapeHtml(filter.namaUnit))
                .append(" | Shift : ").append(escapeHtml(filter.namaShift))
                .append(" | Keyword : ").append(filter.keyword.isEmpty() ? "-" : escapeHtml(filter.keyword)).append("</font>");
        htmlBuilder.append("</td></tr>");
        htmlBuilder.append("</table><br>");

        htmlBuilder.append("<table width='2300px' border='0' align='center' cellpadding='0' cellspacing='0'>");
        htmlBuilder.append("<tr class='head'>");
        htmlBuilder.append("<td rowspan='2' width='80'>Hari</td>");
        htmlBuilder.append("<td rowspan='2' width='95'>Tanggal</td>");
        htmlBuilder.append("<td rowspan='2' width='150'>Nama Dokter HFIS</td>");
        htmlBuilder.append("<td colspan='2' width='150'>Jam Pelayanan</td>");
        htmlBuilder.append("<td colspan='2' width='120'>Jumlah Pasien</td>");
        htmlBuilder.append("<td colspan='2' width='110'>SKBS</td>");
        htmlBuilder.append("<td colspan='2' width='120'>Tindakan</td>");
        htmlBuilder.append("<td rowspan='2' width='90'>Harga</td>");
        htmlBuilder.append("<td colspan='2' width='120'>USG</td>");
        htmlBuilder.append("<td colspan='3' width='250'>Resep Harga Obat</td>");
        htmlBuilder.append("</tr>");

        htmlBuilder.append("<tr class='head2'>");
        htmlBuilder.append("<td width='75'>Mulai</td>");
        htmlBuilder.append("<td width='75'>Selesai</td>");
        htmlBuilder.append("<td width='60'>BPJS</td>");
        htmlBuilder.append("<td width='60'>UMUM</td>");
        htmlBuilder.append("<td width='55'>BPJS</td>");
        htmlBuilder.append("<td width='55'>UMUM</td>");
        htmlBuilder.append("<td width='60'>BPJS</td>");
        htmlBuilder.append("<td width='60'>UMUM</td>");
        htmlBuilder.append("<td width='60'>BPJS</td>");
        htmlBuilder.append("<td width='60'>UMUM</td>");
        htmlBuilder.append("<td width='80'>BPJS</td>");
        htmlBuilder.append("<td width='80'>Poli Gigi Umum</td>");
        htmlBuilder.append("<td width='90'>Poli Umum Umum</td>");
        htmlBuilder.append("</tr>");

        int jumlahBaris = 0;
        List<BarisLaporan> rows = new ArrayList<BarisLaporan>();
        int totalPasienBpjsLokal = 0;
        int totalPasienUmumLokal = 0;
        double totalSkbsBpjsLokal = 0;
        double totalSkbsUmumLokal = 0;
        double totalTindakanBpjsLokal = 0;
        double totalTindakanUmumLokal = 0;
        double totalHargaLokal = 0;
        double totalUsgBpjsLokal = 0;
        double totalUsgUmumLokal = 0;
        double totalObatBpjsLokal = 0;
        double totalObatGigiUmumLokal = 0;
        double totalObatPoliUmumLokal = 0;
        try {
            String filterDokter = filter.kodeDokter.isEmpty() ? "" : "and rp.kd_dokter=? ";
            String filterPoli = filter.kodePoli.isEmpty() ? "" : "and rp.kd_poli=? ";
            String filterKeyword = filter.keyword.isEmpty() ? "" : "and (d.nm_dokter like ? or p.nm_poli like ? or rp.no_rawat like ? or rp.no_rkm_medis like ?) ";
            String filterShift = filter.shiftAwal.isEmpty() ? "" : "and exists (select 1 from pemeriksaan_ralan prf where prf.no_rawat=rp.no_rawat and prf.jam_rawat>=? and prf.jam_rawat<?) ";
            String filterShiftPeriksa = filter.shiftAwal.isEmpty() ? "" : "and pr.jam_rawat>=? and pr.jam_rawat<? ";
            String sql =
                "select base.tgl_registrasi,base.nm_dokter,base.nm_poli,base.jml_pasien_bpjs,base.jml_pasien_umum," +
                "ifnull(periksa.jam_mulai,'-') as jam_mulai," +
                "ifnull(periksa.jam_selesai,'-') as jam_selesai," +
                "ifnull(trx.skbs_bpjs,0) as skbs_bpjs,ifnull(trx.skbs_umum,0) as skbs_umum," +
                "ifnull(trx.tindakan_bpjs,0) as tindakan_bpjs,ifnull(trx.tindakan_umum,0) as tindakan_umum," +
                "ifnull(trx.harga_tindakan,0) as harga_tindakan," +
                "ifnull(trx.usg_bpjs,0) as usg_bpjs,ifnull(trx.usg_umum,0) as usg_umum," +
                "ifnull(trx.obat_bpjs,0) as obat_bpjs,ifnull(trx.obat_gigi_umum,0) as obat_gigi_umum," +
                "ifnull(trx.obat_poli_umum,0) as obat_poli_umum " +
                "from (" +
                "select rp.tgl_registrasi,rp.kd_dokter,d.nm_dokter,rp.kd_poli,p.nm_poli," +
                "count(rp.no_rawat) as jml_kunjungan," +
                "sum(case when rp.kd_pj='BPJ' and rp.status_bayar='Sudah Bayar' and rp.stts='Sudah' then 1 else 0 end) as jml_pasien_bpjs," +
                "sum(case when rp.kd_pj='A09' and rp.status_bayar='Sudah Bayar' and rp.stts='Sudah' then 1 else 0 end) as jml_pasien_umum " +
                "from reg_periksa rp " +
                "inner join dokter d on d.kd_dokter=rp.kd_dokter " +
                "inner join poliklinik p on p.kd_poli=rp.kd_poli " +
                "where rp.status_lanjut='Ralan' and rp.stts<>'Batal' and rp.tgl_registrasi between ? and ? " +
                filterDokter +
                filterPoli +
                filterKeyword +
                filterShift +
                "group by rp.tgl_registrasi,rp.kd_dokter,d.nm_dokter,rp.kd_poli,p.nm_poli " +
                "having (sum(case when rp.kd_pj='BPJ' and rp.status_bayar='Sudah Bayar' and rp.stts='Sudah' then 1 else 0 end) + " +
                "sum(case when rp.kd_pj='A09' and rp.status_bayar='Sudah Bayar' and rp.stts='Sudah' then 1 else 0 end)) > 0" +
                ") as base " +
                "left join (" +
                "select rp.tgl_registrasi,rp.kd_dokter,rp.kd_poli," +
                "date_format(min(concat(pr.tgl_perawatan,' ',pr.jam_rawat)),'%H:%i') as jam_mulai," +
                "date_format(max(concat(pr.tgl_perawatan,' ',pr.jam_rawat)),'%H:%i') as jam_selesai " +
                "from reg_periksa rp " +
                "inner join dokter d on d.kd_dokter=rp.kd_dokter " +
                "inner join poliklinik p on p.kd_poli=rp.kd_poli " +
                "inner join pemeriksaan_ralan pr on pr.no_rawat=rp.no_rawat " +
                "where rp.status_lanjut='Ralan' and rp.stts<>'Batal' and rp.tgl_registrasi between ? and ? " +
                filterDokter +
                filterPoli +
                filterKeyword +
                filterShiftPeriksa +
                "group by rp.tgl_registrasi,rp.kd_dokter,rp.kd_poli" +
                ") as periksa on periksa.tgl_registrasi=base.tgl_registrasi and periksa.kd_dokter=base.kd_dokter and periksa.kd_poli=base.kd_poli " +
                "left join (" +
                "select rp.tgl_registrasi,rp.kd_dokter,rp.kd_poli," +
                "sum(case when rp.kd_pj='BPJ' and lower(ifnull(tag.nama_item,'')) like '%skbs%' then tag.qty else 0 end) as skbs_bpjs," +
                "sum(case when rp.kd_pj='A09' and lower(ifnull(tag.nama_item,'')) like '%skbs%' then tag.qty else 0 end) as skbs_umum," +
                "sum(case when rp.kd_pj='BPJ' and lower(ifnull(tag.nama_item,'')) like '%usg%' then tag.qty else 0 end) as usg_bpjs," +
                "sum(case when rp.kd_pj='A09' and lower(ifnull(tag.nama_item,'')) like '%usg%' then tag.qty else 0 end) as usg_umum," +
                "sum(case when rp.kd_pj='BPJ' and lower(ifnull(tag.nama_item,'')) not like '%skbs%' and lower(ifnull(tag.nama_item,'')) not like '%usg%' and lower(ifnull(tag.nama_item,'')) not like '%obat%' and lower(ifnull(tag.nama_item,'')) not like '%resep%' and lower(ifnull(tag.nama_item,'')) not like '%farmasi%' then tag.qty else 0 end) as tindakan_bpjs," +
                "sum(case when rp.kd_pj='A09' and lower(ifnull(tag.nama_item,'')) not like '%skbs%' and lower(ifnull(tag.nama_item,'')) not like '%usg%' and lower(ifnull(tag.nama_item,'')) not like '%obat%' and lower(ifnull(tag.nama_item,'')) not like '%resep%' and lower(ifnull(tag.nama_item,'')) not like '%farmasi%' then tag.qty else 0 end) as tindakan_umum," +
                "sum(case when lower(ifnull(tag.nama_item,'')) not like '%skbs%' and lower(ifnull(tag.nama_item,'')) not like '%usg%' and lower(ifnull(tag.nama_item,'')) not like '%obat%' and lower(ifnull(tag.nama_item,'')) not like '%resep%' and lower(ifnull(tag.nama_item,'')) not like '%farmasi%' then tag.nilai else 0 end) as harga_tindakan," +
                "sum(case when rp.kd_pj='BPJ' and (tag.status_item='Obat' or lower(ifnull(tag.nama_item,'')) like '%obat%' or lower(ifnull(tag.nama_item,'')) like '%resep%' or lower(ifnull(tag.nama_item,'')) like '%farmasi%') then tag.nilai else 0 end) as obat_bpjs," +
                "sum(case when rp.kd_pj='A09' and lower(p.nm_poli) like '%gigi%' and (tag.status_item='Obat' or lower(ifnull(tag.nama_item,'')) like '%obat%' or lower(ifnull(tag.nama_item,'')) like '%resep%' or lower(ifnull(tag.nama_item,'')) like '%farmasi%') then tag.nilai else 0 end) as obat_gigi_umum," +
                "sum(case when rp.kd_pj='A09' and lower(p.nm_poli) like '%umum%' and (tag.status_item='Obat' or lower(ifnull(tag.nama_item,'')) like '%obat%' or lower(ifnull(tag.nama_item,'')) like '%resep%' or lower(ifnull(tag.nama_item,'')) like '%farmasi%') then tag.nilai else 0 end) as obat_poli_umum " +
                "from reg_periksa rp " +
                "inner join dokter d on d.kd_dokter=rp.kd_dokter " +
                "inner join poliklinik p on p.kd_poli=rp.kd_poli " +
                "inner join (" +
                "select billing.no_rawat,billing.nm_perawatan as nama_item,billing.status as status_item,ifnull(billing.jumlah,1) as qty,billing.totalbiaya as nilai,'billing' as sumber from billing " +
                "union all " +
                "select tambahan_biaya.no_rawat,tambahan_biaya.nama_biaya as nama_item,'Tambahan' as status_item,1 as qty,tambahan_biaya.besar_biaya as nilai,'tambahan' as sumber from tambahan_biaya" +
                ") as tag on tag.no_rawat=rp.no_rawat " +
                "where rp.status_lanjut='Ralan' and rp.stts='Sudah' and rp.status_bayar='Sudah Bayar' and rp.tgl_registrasi between ? and ? " +
                filterDokter +
                filterPoli +
                filterKeyword +
                filterShift +
                "group by rp.tgl_registrasi,rp.kd_dokter,rp.kd_poli" +
                ") as trx on trx.tgl_registrasi=base.tgl_registrasi and trx.kd_dokter=base.kd_dokter and trx.kd_poli=base.kd_poli " +
                "group by base.tgl_registrasi,base.kd_dokter,base.nm_dokter,base.kd_poli,base.nm_poli,base.jml_pasien_bpjs,base.jml_pasien_umum,periksa.jam_mulai,periksa.jam_selesai,trx.skbs_bpjs,trx.skbs_umum,trx.tindakan_bpjs,trx.tindakan_umum,trx.harga_tindakan,trx.usg_bpjs,trx.usg_umum,trx.obat_bpjs,trx.obat_gigi_umum,trx.obat_poli_umum " +
                "order by base.tgl_registrasi,base.nm_dokter,base.nm_poli";
            ps = koneksi.prepareStatement(sql);
            int pIndex = 1;
            ps.setString(pIndex++, tanggalAwalSql);
            ps.setString(pIndex++, tanggalAkhirSql);
            if (!filter.kodeDokter.isEmpty()) {
                ps.setString(pIndex++, filter.kodeDokter);
            }
            if (!filter.kodePoli.isEmpty()) {
                ps.setString(pIndex++, filter.kodePoli);
            }
            if (!filter.keyword.isEmpty()) {
                for (int i = 0; i < 4; i++) {
                    ps.setString(pIndex++, "%" + filter.keyword + "%");
                }
            }
            if (!filter.shiftAwal.isEmpty()) {
                ps.setString(pIndex++, filter.shiftAwal);
                ps.setString(pIndex++, filter.shiftAkhir);
            }
            ps.setString(pIndex++, tanggalAwalSql);
            ps.setString(pIndex++, tanggalAkhirSql);
            if (!filter.kodeDokter.isEmpty()) {
                ps.setString(pIndex++, filter.kodeDokter);
            }
            if (!filter.kodePoli.isEmpty()) {
                ps.setString(pIndex++, filter.kodePoli);
            }
            if (!filter.keyword.isEmpty()) {
                for (int i = 0; i < 4; i++) {
                    ps.setString(pIndex++, "%" + filter.keyword + "%");
                }
            }
            if (!filter.shiftAwal.isEmpty()) {
                ps.setString(pIndex++, filter.shiftAwal);
                ps.setString(pIndex++, filter.shiftAkhir);
            }
            ps.setString(pIndex++, tanggalAwalSql);
            ps.setString(pIndex++, tanggalAkhirSql);
            if (!filter.kodeDokter.isEmpty()) {
                ps.setString(pIndex++, filter.kodeDokter);
            }
            if (!filter.kodePoli.isEmpty()) {
                ps.setString(pIndex++, filter.kodePoli);
            }
            if (!filter.keyword.isEmpty()) {
                for (int i = 0; i < 4; i++) {
                    ps.setString(pIndex++, "%" + filter.keyword + "%");
                }
            }
            if (!filter.shiftAwal.isEmpty()) {
                ps.setString(pIndex++, filter.shiftAwal);
                ps.setString(pIndex++, filter.shiftAkhir);
            }
            rs = ps.executeQuery();
            while (rs.next()) {
                jumlahBaris++;
                LocalDate tanggal = rs.getDate("tgl_registrasi").toLocalDate();
                int jmlPasienBpjs = rs.getInt("jml_pasien_bpjs");
                int jmlPasienUmum = rs.getInt("jml_pasien_umum");
                double skbsBpjs = rs.getDouble("skbs_bpjs");
                double skbsUmum = rs.getDouble("skbs_umum");
                double tindakanBpjs = rs.getDouble("tindakan_bpjs");
                double tindakanUmum = rs.getDouble("tindakan_umum");
                double hargaTindakan = rs.getDouble("harga_tindakan");
                double usgBpjs = rs.getDouble("usg_bpjs");
                double usgUmum = rs.getDouble("usg_umum");
                double obatBpjs = rs.getDouble("obat_bpjs");
                double obatGigiUmum = rs.getDouble("obat_gigi_umum");
                double obatPoliUmum = rs.getDouble("obat_poli_umum");
                totalPasienBpjsLokal += jmlPasienBpjs;
                totalPasienUmumLokal += jmlPasienUmum;
                totalSkbsBpjsLokal += skbsBpjs;
                totalSkbsUmumLokal += skbsUmum;
                totalTindakanBpjsLokal += tindakanBpjs;
                totalTindakanUmumLokal += tindakanUmum;
                totalHargaLokal += hargaTindakan;
                totalUsgBpjsLokal += usgBpjs;
                totalUsgUmumLokal += usgUmum;
                totalObatBpjsLokal += obatBpjs;
                totalObatGigiUmumLokal += obatGigiUmum;
                totalObatPoliUmumLokal += obatPoliUmum;
                rows.add(new BarisLaporan(
                    namaHari(tanggal.getDayOfWeek()),
                    tampilFormat.format(tanggal),
                    rs.getString("nm_dokter"),
                    rs.getString("nm_poli"),
                    rs.getString("jam_mulai"),
                    rs.getString("jam_selesai"),
                    jmlPasienBpjs,
                    jmlPasienUmum,
                    skbsBpjs,
                    skbsUmum,
                    tindakanBpjs,
                    tindakanUmum,
                    hargaTindakan,
                    usgBpjs,
                    usgUmum,
                    obatBpjs,
                    obatGigiUmum,
                    obatPoliUmum
                ));
                htmlBuilder.append("<tr class='isi'>");
                htmlBuilder.append("<td class='tengah'>").append(namaHari(tanggal.getDayOfWeek())).append("</td>");
                htmlBuilder.append("<td class='tengah'>").append(tampilFormat.format(tanggal)).append("</td>");
                htmlBuilder.append("<td>").append(escapeHtml(rs.getString("nm_dokter"))).append("<br><span class='kosong'>").append(escapeHtml(rs.getString("nm_poli"))).append("</span></td>");
                htmlBuilder.append("<td class='tengah'>").append(rs.getString("jam_mulai")).append("</td>");
                htmlBuilder.append("<td class='tengah'>").append(rs.getString("jam_selesai")).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(jmlPasienBpjs).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(jmlPasienUmum).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(formatJumlah(skbsBpjs)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(formatJumlah(skbsUmum)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(formatJumlah(tindakanBpjs)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(formatJumlah(tindakanUmum)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(Valid.SetAngka(hargaTindakan)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(formatJumlah(usgBpjs)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(formatJumlah(usgUmum)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(Valid.SetAngka(obatBpjs)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(Valid.SetAngka(obatGigiUmum)).append("</td>");
                htmlBuilder.append("<td class='kanan'>").append(Valid.SetAngka(obatPoliUmum)).append("</td>");
                htmlBuilder.append("</tr>");
            }
        } catch (Exception e) {
            htmlBuilder.append("<tr class='isi'><td colspan='17' align='center'>Gagal menampilkan data : ").append(escapeHtml(e.getMessage())).append("</td></tr>");
            System.out.println("Notifikasi : " + e);
        } finally {
            try {
                if (rs != null) {
                    rs.close();
                }
                if (ps != null) {
                    ps.close();
                }
            } catch (Exception e) {
                System.out.println("Notifikasi : " + e);
            }
        }

        if (jumlahBaris == 0) {
            htmlBuilder.append("<tr class='isi'><td colspan='17' align='center'>Tidak ada data pelayanan yang sesuai dengan filter.</td></tr>");
        }

        htmlBuilder.append("<tr class='head2'>");
        htmlBuilder.append("<td colspan='5' align='right'>Total Periode</td>");
        htmlBuilder.append("<td class='kanan'>").append(totalPasienBpjsLokal).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(totalPasienUmumLokal).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(formatJumlah(totalSkbsBpjsLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(formatJumlah(totalSkbsUmumLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(formatJumlah(totalTindakanBpjsLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(formatJumlah(totalTindakanUmumLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(Valid.SetAngka(totalHargaLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(formatJumlah(totalUsgBpjsLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(formatJumlah(totalUsgUmumLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(Valid.SetAngka(totalObatBpjsLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(Valid.SetAngka(totalObatGigiUmumLokal)).append("</td>");
        htmlBuilder.append("<td class='kanan'>").append(Valid.SetAngka(totalObatPoliUmumLokal)).append("</td>");
        htmlBuilder.append("</tr>");

        htmlBuilder.append("</table>");
        htmlBuilder.append("</body></html>");
        return new ReportSnapshot(
            htmlBuilder.toString(), rows,
            totalPasienBpjsLokal, totalPasienUmumLokal,
            totalSkbsBpjsLokal, totalSkbsUmumLokal,
            totalTindakanBpjsLokal, totalTindakanUmumLokal,
            totalHargaLokal, totalUsgBpjsLokal, totalUsgUmumLokal,
            totalObatBpjsLokal, totalObatGigiUmumLokal, totalObatPoliUmumLokal
        );
    }

    private void loadDokter() {
        loadPilihan(
            cmbDokter,
            "Semua Dokter",
            "select kd_dokter,nm_dokter from dokter where status='1' order by nm_dokter",
            "kd_dokter",
            "nm_dokter"
        );
    }

    private void loadUnit() {
        loadPilihan(
            cmbUnit,
            "Semua Unit",
            "select kd_poli,nm_poli from poliklinik order by nm_poli",
            "kd_poli",
            "nm_poli"
        );
    }

    private void loadPilihan(JComboBox combo, String labelSemua, String sql, String kolomKode, String kolomNama) {
        DefaultComboBoxModel model = new DefaultComboBoxModel();
        model.addElement(labelSemua);
        PreparedStatement psLocal = null;
        ResultSet rsLocal = null;
        try {
            psLocal = koneksi.prepareStatement(sql);
            rsLocal = psLocal.executeQuery();
            while (rsLocal.next()) {
                model.addElement(rsLocal.getString(kolomKode) + " - " + rsLocal.getString(kolomNama));
            }
        } catch (Exception e) {
            System.out.println("Notifikasi : " + e);
        } finally {
            try {
                if (rsLocal != null) {
                    rsLocal.close();
                }
                if (psLocal != null) {
                    psLocal.close();
                }
            } catch (Exception e) {
                System.out.println("Notifikasi : " + e);
            }
        }
        combo.setModel(model);
    }

    private String namaHari(DayOfWeek dayOfWeek) {
        switch (dayOfWeek) {
            case MONDAY:
                return "SENIN";
            case TUESDAY:
                return "SELASA";
            case WEDNESDAY:
                return "RABU";
            case THURSDAY:
                return "KAMIS";
            case FRIDAY:
                return "JUMAT";
            case SATURDAY:
                return "SABTU";
            default:
                return "MINGGU";
        }
    }

    private String getKodePilihan(JComboBox combo, String labelSemua) {
        String pilihan = combo.getSelectedItem() == null ? "" : combo.getSelectedItem().toString();
        if (pilihan.equals(labelSemua) || !pilihan.contains(" - ")) {
            return "";
        }
        return pilihan.substring(0, pilihan.indexOf(" - ")).trim();
    }

    private String formatJumlah(double nilai) {
        if (Math.abs(nilai - Math.rint(nilai)) < 0.0001) {
            return String.valueOf((long) Math.rint(nilai));
        }
        return String.format(Locale.US, "%.2f", nilai);
    }

    private String escapeHtml(String text) {
        if (text == null) {
            return "";
        }
        return text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");
    }

    private String getNamaShift() {
        return cmbShift.getSelectedItem() == null ? "Semua Shift" : cmbShift.getSelectedItem().toString();
    }

    private String getShiftAwal() {
        if (cmbShift.getSelectedIndex() == 1) {
            return "08:00:00";
        } else if (cmbShift.getSelectedIndex() == 2) {
            return "15:00:00";
        }
        return "";
    }

    private String getShiftAkhir() {
        if (cmbShift.getSelectedIndex() == 1) {
            return "15:00:00";
        } else if (cmbShift.getSelectedIndex() == 2) {
            return "22:00:01";
        }
        return "";
    }

    private FilterLaporan ambilFilterSaatIni() {
        return new FilterLaporan(
            Valid.SetTgl(Tgl1.getSelectedItem() + ""),
            Valid.SetTgl(Tgl2.getSelectedItem() + ""),
            getKodePilihan(cmbDokter, "Semua Dokter"),
            getKodePilihan(cmbUnit, "Semua Unit"),
            cmbDokter.getSelectedItem() == null ? "Semua Dokter" : cmbDokter.getSelectedItem().toString(),
            cmbUnit.getSelectedItem() == null ? "Semua Unit" : cmbUnit.getSelectedItem().toString(),
            getNamaShift(),
            getShiftAwal(),
            getShiftAkhir(),
            TCari.getText().trim()
        );
    }

    private void terapkanSnapshot(ReportSnapshot snapshot) {
        htmlContent = new StringBuilder(snapshot.html);
        dataLaporan.clear();
        dataLaporan.addAll(snapshot.rows);
        totalPasienBpjs = snapshot.totalPasienBpjs;
        totalPasienUmum = snapshot.totalPasienUmum;
        totalSkbsBpjs = snapshot.totalSkbsBpjs;
        totalSkbsUmum = snapshot.totalSkbsUmum;
        totalTindakanBpjs = snapshot.totalTindakanBpjs;
        totalTindakanUmum = snapshot.totalTindakanUmum;
        totalHarga = snapshot.totalHarga;
        totalUsgBpjs = snapshot.totalUsgBpjs;
        totalUsgUmum = snapshot.totalUsgUmum;
        totalObatBpjs = snapshot.totalObatBpjs;
        totalObatGigiUmum = snapshot.totalObatGigiUmum;
        totalObatPoliUmum = snapshot.totalObatPoliUmum;
        LoadHTML.setText(snapshot.html);
        LoadHTML.setCaretPosition(0);
    }

    private void setLoadingState(boolean loading, String pesan) {
        sedangMemuat = loading;
        this.setCursor(loading ? Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR) : Cursor.getDefaultCursor());
        labelInfo.setText(pesan);
        btnTampilkan.setEnabled(!loading);
        btnPeriodeBulan.setEnabled(!loading);
        BtnExport.setEnabled(!loading && !dataLaporan.isEmpty());
        BtnPrint.setEnabled(!loading);
    }

    private void exportKeExcel() {
        if (dataLaporan.isEmpty()) {
            JOptionPane.showMessageDialog(rootPane, "Data laporan belum tersedia untuk diexport.");
            return;
        }

        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Simpan Laporan Excel");
        chooser.setSelectedFile(new File("LaporanKegiatanPelayanan_" +
            Valid.SetTgl(Tgl1.getSelectedItem() + "") + "_" +
            Valid.SetTgl(Tgl2.getSelectedItem() + "") + ".xls"));
        chooser.setFileFilter(new FileNameExtensionFilter("Excel 97-2003 (*.xls)", "xls"));

        if (chooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File file = chooser.getSelectedFile();
        if (!file.getName().toLowerCase().endsWith(".xls")) {
            file = new File(file.getAbsolutePath() + ".xls");
        }

        this.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        WritableWorkbook workbook = null;
        try {
            workbook = Workbook.createWorkbook(file);
            WritableSheet sheet = workbook.createSheet("Laporan", 0);
            int row = 0;

            sheet.addCell(new jxl.write.Label(0, row++, akses.getnamars()));
            sheet.addCell(new jxl.write.Label(0, row++, "Laporan Kegiatan Selama Pelayanan"));
            sheet.addCell(new jxl.write.Label(0, row++, "Periode " +
                Valid.SetTgl(Tgl1.getSelectedItem() + "") + " s.d. " +
                Valid.SetTgl(Tgl2.getSelectedItem() + "")));
            sheet.addCell(new jxl.write.Label(0, row++, "Dokter: " + cmbDokter.getSelectedItem() +
                " | Unit: " + cmbUnit.getSelectedItem() +
                " | Shift: " + cmbShift.getSelectedItem() +
                " | Keyword: " + (TCari.getText().trim().isEmpty() ? "-" : TCari.getText().trim())));
            row++;

            String[] header = {
                "Hari", "Tanggal", "Nama Dokter", "Unit", "Jam Mulai", "Jam Selesai",
                "Pasien BPJS", "Pasien Umum", "SKBS BPJS", "SKBS Umum",
                "Tindakan BPJS", "Tindakan Umum", "Harga",
                "USG BPJS", "USG Umum", "Obat BPJS", "Obat Poli Gigi Umum", "Obat Poli Umum Umum"
            };

            for (int i = 0; i < header.length; i++) {
                sheet.addCell(new jxl.write.Label(i, row, header[i]));
            }
            row++;

            for (BarisLaporan item : dataLaporan) {
                sheet.addCell(new jxl.write.Label(0, row, item.hari));
                sheet.addCell(new jxl.write.Label(1, row, item.tanggal));
                sheet.addCell(new jxl.write.Label(2, row, item.namaDokter));
                sheet.addCell(new jxl.write.Label(3, row, item.namaPoli));
                sheet.addCell(new jxl.write.Label(4, row, item.jamMulai));
                sheet.addCell(new jxl.write.Label(5, row, item.jamSelesai));
                sheet.addCell(new jxl.write.Number(6, row, item.jumlahPasienBpjs));
                sheet.addCell(new jxl.write.Number(7, row, item.jumlahPasienUmum));
                sheet.addCell(new jxl.write.Number(8, row, item.skbsBpjs));
                sheet.addCell(new jxl.write.Number(9, row, item.skbsUmum));
                sheet.addCell(new jxl.write.Number(10, row, item.tindakanBpjs));
                sheet.addCell(new jxl.write.Number(11, row, item.tindakanUmum));
                sheet.addCell(new jxl.write.Number(12, row, item.hargaTindakan));
                sheet.addCell(new jxl.write.Number(13, row, item.usgBpjs));
                sheet.addCell(new jxl.write.Number(14, row, item.usgUmum));
                sheet.addCell(new jxl.write.Number(15, row, item.obatBpjs));
                sheet.addCell(new jxl.write.Number(16, row, item.obatGigiUmum));
                sheet.addCell(new jxl.write.Number(17, row, item.obatPoliUmum));
                row++;
            }

            sheet.addCell(new jxl.write.Label(0, row, "Total Periode"));
            sheet.addCell(new jxl.write.Number(6, row, totalPasienBpjs));
            sheet.addCell(new jxl.write.Number(7, row, totalPasienUmum));
            sheet.addCell(new jxl.write.Number(8, row, totalSkbsBpjs));
            sheet.addCell(new jxl.write.Number(9, row, totalSkbsUmum));
            sheet.addCell(new jxl.write.Number(10, row, totalTindakanBpjs));
            sheet.addCell(new jxl.write.Number(11, row, totalTindakanUmum));
            sheet.addCell(new jxl.write.Number(12, row, totalHarga));
            sheet.addCell(new jxl.write.Number(13, row, totalUsgBpjs));
            sheet.addCell(new jxl.write.Number(14, row, totalUsgUmum));
            sheet.addCell(new jxl.write.Number(15, row, totalObatBpjs));
            sheet.addCell(new jxl.write.Number(16, row, totalObatGigiUmum));
            sheet.addCell(new jxl.write.Number(17, row, totalObatPoliUmum));

            for (int i = 0; i < header.length; i++) {
                sheet.setColumnView(i, 18);
            }
            sheet.setColumnView(2, 28);
            sheet.setColumnView(3, 24);

            workbook.write();
            JOptionPane.showMessageDialog(rootPane, "File Excel berhasil dibuat:\n" + file.getAbsolutePath());
        } catch (Exception e) {
            JOptionPane.showMessageDialog(rootPane, "Export Excel gagal: " + e.getMessage());
            System.out.println("Notifikasi : " + e);
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (Exception e) {
                System.out.println("Notifikasi : " + e);
            }
            this.setCursor(Cursor.getDefaultCursor());
        }
    }

    private void runOnUi(Runnable task) {
        SwingUtilities.invokeLater(task);
    }

    private static class FilterLaporan {
        private final String tanggalAwal;
        private final String tanggalAkhir;
        private final String kodeDokter;
        private final String kodePoli;
        private final String namaDokter;
        private final String namaUnit;
        private final String namaShift;
        private final String shiftAwal;
        private final String shiftAkhir;
        private final String keyword;

        private FilterLaporan(String tanggalAwal, String tanggalAkhir, String kodeDokter, String kodePoli,
                String namaDokter, String namaUnit, String namaShift, String shiftAwal, String shiftAkhir,
                String keyword) {
            this.tanggalAwal = tanggalAwal;
            this.tanggalAkhir = tanggalAkhir;
            this.kodeDokter = kodeDokter;
            this.kodePoli = kodePoli;
            this.namaDokter = namaDokter;
            this.namaUnit = namaUnit;
            this.namaShift = namaShift;
            this.shiftAwal = shiftAwal;
            this.shiftAkhir = shiftAkhir;
            this.keyword = keyword;
        }
    }

    private static class ReportSnapshot {
        private final String html;
        private final List<BarisLaporan> rows;
        private final int totalPasienBpjs;
        private final int totalPasienUmum;
        private final double totalSkbsBpjs;
        private final double totalSkbsUmum;
        private final double totalTindakanBpjs;
        private final double totalTindakanUmum;
        private final double totalHarga;
        private final double totalUsgBpjs;
        private final double totalUsgUmum;
        private final double totalObatBpjs;
        private final double totalObatGigiUmum;
        private final double totalObatPoliUmum;

        private ReportSnapshot(String html, List<BarisLaporan> rows, int totalPasienBpjs, int totalPasienUmum,
                double totalSkbsBpjs, double totalSkbsUmum, double totalTindakanBpjs, double totalTindakanUmum,
                double totalHarga, double totalUsgBpjs, double totalUsgUmum, double totalObatBpjs,
                double totalObatGigiUmum, double totalObatPoliUmum) {
            this.html = html;
            this.rows = rows;
            this.totalPasienBpjs = totalPasienBpjs;
            this.totalPasienUmum = totalPasienUmum;
            this.totalSkbsBpjs = totalSkbsBpjs;
            this.totalSkbsUmum = totalSkbsUmum;
            this.totalTindakanBpjs = totalTindakanBpjs;
            this.totalTindakanUmum = totalTindakanUmum;
            this.totalHarga = totalHarga;
            this.totalUsgBpjs = totalUsgBpjs;
            this.totalUsgUmum = totalUsgUmum;
            this.totalObatBpjs = totalObatBpjs;
            this.totalObatGigiUmum = totalObatGigiUmum;
            this.totalObatPoliUmum = totalObatPoliUmum;
        }
    }

    private static class BarisLaporan {
        private final String hari;
        private final String tanggal;
        private final String namaDokter;
        private final String namaPoli;
        private final String jamMulai;
        private final String jamSelesai;
        private final int jumlahPasienBpjs;
        private final int jumlahPasienUmum;
        private final double skbsBpjs;
        private final double skbsUmum;
        private final double tindakanBpjs;
        private final double tindakanUmum;
        private final double hargaTindakan;
        private final double usgBpjs;
        private final double usgUmum;
        private final double obatBpjs;
        private final double obatGigiUmum;
        private final double obatPoliUmum;

        private BarisLaporan(String hari, String tanggal, String namaDokter, String namaPoli,
                String jamMulai, String jamSelesai, int jumlahPasienBpjs, int jumlahPasienUmum,
                double skbsBpjs, double skbsUmum, double tindakanBpjs, double tindakanUmum,
                double hargaTindakan, double usgBpjs, double usgUmum, double obatBpjs,
                double obatGigiUmum, double obatPoliUmum) {
            this.hari = hari;
            this.tanggal = tanggal;
            this.namaDokter = namaDokter;
            this.namaPoli = namaPoli;
            this.jamMulai = jamMulai;
            this.jamSelesai = jamSelesai;
            this.jumlahPasienBpjs = jumlahPasienBpjs;
            this.jumlahPasienUmum = jumlahPasienUmum;
            this.skbsBpjs = skbsBpjs;
            this.skbsUmum = skbsUmum;
            this.tindakanBpjs = tindakanBpjs;
            this.tindakanUmum = tindakanUmum;
            this.hargaTindakan = hargaTindakan;
            this.usgBpjs = usgBpjs;
            this.usgUmum = usgUmum;
            this.obatBpjs = obatBpjs;
            this.obatGigiUmum = obatGigiUmum;
            this.obatPoliUmum = obatPoliUmum;
        }
    }

    public static void main(String args[]) {
        java.awt.EventQueue.invokeLater(() -> {
            DlgLaporanKegiatanPelayanan dialog = new DlgLaporanKegiatanPelayanan(new javax.swing.JFrame(), true);
            dialog.addWindowListener(new java.awt.event.WindowAdapter() {
                @Override
                public void windowClosing(java.awt.event.WindowEvent e) {
                    System.exit(0);
                }
            });
            dialog.setVisible(true);
        });
    }

    private widget.Button BtnExport;
    private widget.Button BtnKeluar;
    private widget.Button BtnPrint;
    private widget.editorpane LoadHTML;
    private widget.ScrollPane Scroll;
    private widget.TextBox TCari;
    private widget.Tanggal Tgl1;
    private widget.Tanggal Tgl2;
    private widget.Button btnPeriodeBulan;
    private widget.Button btnTampilkan;
    private widget.ComboBox cmbDokter;
    private widget.ComboBox cmbShift;
    private widget.ComboBox cmbUnit;
    private widget.InternalFrame internalFrame1;
    private widget.Label label11;
    private widget.Label label18;
    private widget.Label labelDokter;
    private widget.Label labelInfo;
    private widget.Label labelKeyword;
    private widget.Label labelShift;
    private widget.Label labelUnit;
    private widget.panelisi panelAksi;
    private widget.panelisi panelFilter;
    private widget.panelisi panelFilterInput;
}
