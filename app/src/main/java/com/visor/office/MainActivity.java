package com.visor.office;

import android.app.Activity;
import android.content.Intent;
import android.database.Cursor;
import android.graphics.Bitmap;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.Typeface;
import android.graphics.drawable.ColorDrawable;
import android.graphics.pdf.PdfRenderer;
import android.net.Uri;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.os.ParcelFileDescriptor;
import android.provider.OpenableColumns;
import android.text.SpannableStringBuilder;
import android.text.Spanned;
import android.text.style.RelativeSizeSpan;
import android.text.style.StyleSpan;
import android.util.LruCache;
import android.util.TypedValue;
import android.util.Xml;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.widget.AbsListView;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.BaseAdapter;
import android.widget.Button;
import android.widget.FrameLayout;
import android.widget.HorizontalScrollView;
import android.widget.ImageView;
import android.widget.LinearLayout;
import android.widget.ListView;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import org.xmlpull.v1.XmlPullParser;

import java.io.BufferedInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.MathContext;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/** Visor liviano y 100% offline de PDF, Word (.docx) y Excel (.xlsx). Sin librerias externas. */
public class MainActivity extends Activity {
    static final int REQ_OPEN = 1;
    static final int MAX_ROWS = 3000, MAX_COLS = 40, MAX_COLW = 28;
    static final String MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    static final String MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    private FrameLayout content;
    private TextView title;
    private Spinner sheetSpinner;
    private final Handler ui = new Handler(Looper.getMainLooper());
    private PdfRenderer renderer;
    private ParcelFileDescriptor pfd;
    private int loadToken = 0;
    private float textSp = 17f;
    private Runnable refresh;
    private Uri currentUri;
    private List<String> sheetPaths = new ArrayList<>();
    private List<String> shared = new ArrayList<>();
    private List<String[]> rows = new ArrayList<>();

    static class SheetData { List<String[]> rows = new ArrayList<>(); boolean truncated; }

    // ------------------------------------------------------------ UI base
    @Override protected void onCreate(Bundle b) {
        super.onCreate(b);
        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);

        LinearLayout bar = new LinearLayout(this);
        bar.setGravity(Gravity.CENTER_VERTICAL);
        bar.setBackgroundColor(0xFFE0E0E0);
        bar.setPadding(dp(6), dp(2), dp(6), dp(2));

        Button open = new Button(this);
        open.setText("Abrir archivo");
        open.setAllCaps(false);
        open.setOnClickListener(v -> pick());

        title = new TextView(this);
        title.setSingleLine(true);
        title.setEllipsize(android.text.TextUtils.TruncateAt.END);
        title.setPadding(dp(10), 0, dp(10), 0);

        sheetSpinner = new Spinner(this);
        sheetSpinner.setVisibility(View.GONE);

        Button minus = smallButton("A-", -2f);
        Button plus = smallButton("A+", 2f);

        bar.addView(open);
        bar.addView(title, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1f));
        bar.addView(sheetSpinner);
        bar.addView(minus);
        bar.addView(plus);

        content = new FrameLayout(this);
        root.addView(bar, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.WRAP_CONTENT));
        root.addView(content, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, 0, 1f));
        setContentView(root);

        showMessage("Toc\u00e1 \u00abAbrir archivo\u00bb para elegir un PDF, Word (.docx) o Excel (.xlsx).");
        handleIntent(getIntent());
    }

    private Button smallButton(String text, final float delta) {
        Button bt = new Button(this);
        bt.setText(text);
        bt.setAllCaps(false);
        bt.setMinWidth(dp(48));
        bt.setMinimumWidth(dp(48));
        bt.setOnClickListener(v -> {
            textSp = Math.max(10f, Math.min(34f, textSp + delta));
            if (refresh != null) refresh.run();
        });
        return bt;
    }

    private int dp(float v) { return (int) (v * getResources().getDisplayMetrics().density + 0.5f); }

    private void showMessage(String msg) {
        content.removeAllViews();
        TextView t = new TextView(this);
        t.setText(msg);
        t.setTextSize(18);
        t.setPadding(dp(20), dp(20), dp(20), dp(20));
        content.addView(t);
    }

    private void pick() {
        Intent i = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        i.addCategory(Intent.CATEGORY_OPENABLE);
        i.setType("*/*");
        i.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{"application/pdf", MIME_DOCX, MIME_XLSX});
        startActivityForResult(i, REQ_OPEN);
    }

    @Override protected void onActivityResult(int req, int res, Intent data) {
        super.onActivityResult(req, res, data);
        if (req == REQ_OPEN && res == RESULT_OK && data != null && data.getData() != null) openFile(data.getData());
    }

    @Override protected void onNewIntent(Intent intent) { super.onNewIntent(intent); setIntent(intent); handleIntent(intent); }

    private void handleIntent(Intent i) {
        if (i != null && Intent.ACTION_VIEW.equals(i.getAction()) && i.getData() != null) openFile(i.getData());
    }

    @Override protected void onDestroy() { closePdf(); super.onDestroy(); }

    private String displayName(Uri uri) {
        try (Cursor c = getContentResolver().query(uri, null, null, null, null)) {
            if (c != null && c.moveToFirst()) {
                int ix = c.getColumnIndex(OpenableColumns.DISPLAY_NAME);
                if (ix >= 0) return c.getString(ix);
            }
        } catch (Exception ignored) { }
        String p = uri.getLastPathSegment();
        return p == null ? "archivo" : p;
    }

    private void openFile(Uri uri) {
        closePdf();
        refresh = null;
        sheetSpinner.setVisibility(View.GONE);
        String name = displayName(uri);
        title.setText(name);
        String ext = name.toLowerCase(Locale.ROOT);
        int dot = ext.lastIndexOf('.');
        ext = dot >= 0 ? ext.substring(dot + 1) : "";
        String type = getContentResolver().getType(uri);
        if (type == null) type = "";
        if (ext.equals("pdf") || type.equals("application/pdf")) openPdf(uri);
        else if (ext.equals("docx") || type.contains("wordprocessingml")) openDocx(uri);
        else if (ext.equals("xlsx") || type.contains("spreadsheetml")) openXlsx(uri);
        else if (ext.equals("doc") || ext.equals("xls"))
            showMessage("Los formatos antiguos (.doc / .xls) no est\u00e1n soportados.\nAbrilo en otra app y guardalo como .docx / .xlsx.");
        else showMessage("Formato no soportado. Us\u00e1 PDF, .docx o .xlsx.");
    }

    private static String msg(Throwable e) {
        if (e instanceof OutOfMemoryError) return "el archivo es demasiado grande para la memoria de la tablet";
        return e.getMessage() != null ? e.getMessage() : e.toString();
    }

    // ------------------------------------------------------------ ZIP helper
    private ZipInputStream openEntry(Uri uri, String name) throws IOException {
        InputStream raw = getContentResolver().openInputStream(uri);
        if (raw == null) throw new FileNotFoundException(name);
        ZipInputStream z = new ZipInputStream(new BufferedInputStream(raw, 32768));
        ZipEntry e;
        while ((e = z.getNextEntry()) != null) {
            if (e.getName().equals(name)) return z;
        }
        z.close();
        throw new FileNotFoundException(name);
    }

    // ------------------------------------------------------------ PDF
    private void closePdf() {
        try { if (renderer != null) renderer.close(); } catch (Exception ignored) { }
        try { if (pfd != null) pfd.close(); } catch (Exception ignored) { }
        renderer = null;
        pfd = null;
    }

    private void openPdf(Uri uri) {
        try {
            pfd = getContentResolver().openFileDescriptor(uri, "r");
            renderer = new PdfRenderer(pfd);
        } catch (Exception e) {
            closePdf();
            showMessage("No se pudo abrir el PDF: " + msg(e));
            return;
        }
        final int w = (int) (getResources().getDisplayMetrics().widthPixels * 1.25f);
        final int count = renderer.getPageCount();
        final LruCache<Integer, Bitmap> cache = new LruCache<>(3);
        ListView lv = new ListView(this);
        lv.setBackgroundColor(0xFF777777);
        lv.setDivider(new ColorDrawable(0xFF777777));
        lv.setDividerHeight(dp(8));
        lv.setAdapter(new BaseAdapter() {
            @Override public int getCount() { return count; }
            @Override public Object getItem(int p) { return null; }
            @Override public long getItemId(int p) { return p; }
            @Override public View getView(int pos, View cv, ViewGroup parent) {
                ImageView iv;
                if (cv instanceof ImageView) iv = (ImageView) cv;
                else {
                    iv = new ImageView(MainActivity.this);
                    iv.setAdjustViewBounds(true);
                    iv.setScaleType(ImageView.ScaleType.FIT_CENTER);
                    iv.setLayoutParams(new AbsListView.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.WRAP_CONTENT));
                }
                Bitmap bm = cache.get(pos);
                if (bm == null) {
                    try { bm = renderPage(pos, w); cache.put(pos, bm); } catch (Throwable t) { bm = null; }
                }
                iv.setImageBitmap(bm);
                return iv;
            }
        });
        content.removeAllViews();
        content.addView(lv);
        Toast.makeText(this, count + " p\u00e1ginas", Toast.LENGTH_SHORT).show();
    }

    private Bitmap renderPage(int i, int w) {
        PdfRenderer.Page p = renderer.openPage(i);
        try {
            int h = Math.max(1, Math.round(w * (float) p.getHeight() / p.getWidth()));
            Bitmap b = Bitmap.createBitmap(w, h, Bitmap.Config.ARGB_8888);
            b.eraseColor(Color.WHITE);
            p.render(b, null, null, PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY);
            return b;
        } finally {
            p.close();
        }
    }

    // ------------------------------------------------------------ DOCX
    private void openDocx(final Uri uri) {
        showMessage("Cargando\u2026");
        final int token = ++loadToken;
        new Thread(() -> {
            final List<CharSequence> paras = new ArrayList<>();
            String err = null;
            try (ZipInputStream z = openEntry(uri, "word/document.xml")) {
                parseDocx(z, paras);
            } catch (Throwable e) { err = msg(e); }
            final String fe = err;
            ui.post(() -> {
                if (token != loadToken) return;
                if (fe != null) { showMessage("No se pudo abrir el Word: " + fe); return; }
                if (paras.isEmpty()) { showMessage("El documento no tiene texto."); return; }
                ListView lv = new ListView(this);
                lv.setDividerHeight(0);
                lv.setPadding(dp(16), dp(8), dp(16), dp(8));
                lv.setClipToPadding(false);
                final BaseAdapter ad = new BaseAdapter() {
                    @Override public int getCount() { return paras.size(); }
                    @Override public Object getItem(int p) { return null; }
                    @Override public long getItemId(int p) { return p; }
                    @Override public View getView(int pos, View cv, ViewGroup parent) {
                        TextView t = cv instanceof TextView ? (TextView) cv : new TextView(MainActivity.this);
                        t.setTextSize(TypedValue.COMPLEX_UNIT_SP, textSp);
                        t.setTextColor(0xFF111111);
                        t.setLineSpacing(0f, 1.15f);
                        t.setPadding(0, dp(4), 0, dp(4));
                        CharSequence cs = paras.get(pos);
                        t.setText(cs.length() == 0 ? " " : cs);
                        return t;
                    }
                };
                lv.setAdapter(ad);
                refresh = ad::notifyDataSetChanged;
                content.removeAllViews();
                content.addView(lv);
            });
        }).start();
    }

    private static boolean on(XmlPullParser p) {
        String v = p.getAttributeValue(null, "w:val");
        return v == null || !(v.equals("0") || v.equals("false") || v.equals("off"));
    }

    private void parseDocx(InputStream in, List<CharSequence> out) throws Exception {
        XmlPullParser p = Xml.newPullParser();
        p.setInput(in, null);
        SpannableStringBuilder para = new SpannableStringBuilder();
        SpannableStringBuilder cell = new SpannableStringBuilder();
        SpannableStringBuilder rowB = new SpannableStringBuilder();
        boolean inT = false, inRun = false, bold = false, ital = false, head = false, bullet = false;
        int tbl = 0;
        for (int ev = p.getEventType(); ev != XmlPullParser.END_DOCUMENT; ev = p.next()) {
            if (ev == XmlPullParser.START_TAG) {
                String n = p.getName();
                if (n.equals("w:p")) { para = new SpannableStringBuilder(); head = false; bullet = false; }
                else if (n.equals("w:pStyle")) {
                    String v = p.getAttributeValue(null, "w:val");
                    if (v != null) {
                        v = v.toLowerCase(Locale.ROOT);
                        if (v.contains("heading") || v.contains("ttulo") || v.contains("titulo") || v.equals("title")) head = true;
                    }
                }
                else if (n.equals("w:numPr")) bullet = true;
                else if (n.equals("w:r")) { inRun = true; bold = false; ital = false; }
                else if (n.equals("w:b")) bold = on(p);
                else if (n.equals("w:i")) ital = on(p);
                else if (n.equals("w:t")) inT = true;
                else if (n.equals("w:tab") && inRun) para.append('\t');
                else if ((n.equals("w:br") || n.equals("w:cr")) && inRun) para.append('\n');
                else if (n.equals("w:tbl")) tbl++;
            } else if (ev == XmlPullParser.TEXT) {
                if (inT) {
                    int s = para.length();
                    para.append(p.getText());
                    int e = para.length();
                    if (e > s && bold) para.setSpan(new StyleSpan(Typeface.BOLD), s, e, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
                    if (e > s && ital) para.setSpan(new StyleSpan(Typeface.ITALIC), s, e, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
                }
            } else if (ev == XmlPullParser.END_TAG) {
                String n = p.getName();
                if (n.equals("w:t")) inT = false;
                else if (n.equals("w:r")) inRun = false;
                else if (n.equals("w:p")) {
                    if (bullet) para.insert(0, "\u2022 ");
                    if (head && para.length() > 0) {
                        para.setSpan(new StyleSpan(Typeface.BOLD), 0, para.length(), Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
                        para.setSpan(new RelativeSizeSpan(1.25f), 0, para.length(), Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
                    }
                    if (tbl > 0) cell.append(para).append(' ');
                    else out.add(para);
                }
                else if (n.equals("w:tc")) { rowB.append(cell).append("  \u2502  "); cell = new SpannableStringBuilder(); }
                else if (n.equals("w:tr")) { if (rowB.length() > 0) out.add(rowB); rowB = new SpannableStringBuilder(); }
                else if (n.equals("w:tbl")) { tbl--; if (tbl == 0) out.add(""); }
            }
        }
    }

    // ------------------------------------------------------------ XLSX
    private void openXlsx(final Uri uri) {
        showMessage("Cargando\u2026");
        final int token = ++loadToken;
        new Thread(() -> {
            final List<String> names = new ArrayList<>(), paths = new ArrayList<>(), sst = new ArrayList<>();
            String err = null;
            try {
                readWorkbook(uri, names, paths);
                readShared(uri, sst);
            } catch (Throwable e) { err = msg(e); }
            final String fe = err;
            ui.post(() -> {
                if (token != loadToken) return;
                if (fe != null) { showMessage("No se pudo abrir el Excel: " + fe); return; }
                if (names.isEmpty()) { showMessage("El archivo no tiene hojas."); return; }
                sheetPaths = paths;
                shared = sst;
                currentUri = uri;
                showMessage("Cargando hoja\u2026");
                ArrayAdapter<String> ad = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, names);
                ad.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
                sheetSpinner.setAdapter(ad);
                sheetSpinner.setVisibility(View.VISIBLE);
                sheetSpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
                    @Override public void onItemSelected(AdapterView<?> a, View v, int pos, long id) { loadSheet(pos); }
                    @Override public void onNothingSelected(AdapterView<?> a) { }
                });
            });
        }).start();
    }

    private void readWorkbook(Uri uri, List<String> names, List<String> paths) throws Exception {
        Map<String, String> rels = new HashMap<>();
        try (ZipInputStream z = openEntry(uri, "xl/_rels/workbook.xml.rels")) {
            XmlPullParser p = Xml.newPullParser();
            p.setInput(z, null);
            for (int ev = p.getEventType(); ev != XmlPullParser.END_DOCUMENT; ev = p.next()) {
                if (ev == XmlPullParser.START_TAG && p.getName().equals("Relationship"))
                    rels.put(p.getAttributeValue(null, "Id"), p.getAttributeValue(null, "Target"));
            }
        }
        try (ZipInputStream z = openEntry(uri, "xl/workbook.xml")) {
            XmlPullParser p = Xml.newPullParser();
            p.setInput(z, null);
            for (int ev = p.getEventType(); ev != XmlPullParser.END_DOCUMENT; ev = p.next()) {
                if (ev == XmlPullParser.START_TAG && p.getName().equals("sheet")) {
                    String t = rels.get(p.getAttributeValue(null, "r:id"));
                    if (t == null) continue;
                    names.add(p.getAttributeValue(null, "name"));
                    paths.add(t.startsWith("/") ? t.substring(1) : "xl/" + t);
                }
            }
        }
    }

    private void readShared(Uri uri, List<String> sst) throws Exception {
        ZipInputStream z;
        try { z = openEntry(uri, "xl/sharedStrings.xml"); } catch (FileNotFoundException e) { return; }
        try {
            XmlPullParser p = Xml.newPullParser();
            p.setInput(z, null);
            StringBuilder sb = new StringBuilder();
            boolean inT = false, inPh = false;
            for (int ev = p.getEventType(); ev != XmlPullParser.END_DOCUMENT; ev = p.next()) {
                if (ev == XmlPullParser.START_TAG) {
                    String n = p.getName();
                    if (n.equals("si")) sb.setLength(0);
                    else if (n.equals("rPh")) inPh = true;
                    else if (n.equals("t")) inT = true;
                } else if (ev == XmlPullParser.TEXT) {
                    if (inT && !inPh) sb.append(p.getText());
                } else if (ev == XmlPullParser.END_TAG) {
                    String n = p.getName();
                    if (n.equals("t")) inT = false;
                    else if (n.equals("rPh")) inPh = false;
                    else if (n.equals("si")) sst.add(sb.toString());
                }
            }
        } finally { z.close(); }
    }

    private void loadSheet(final int idx) {
        if (idx < 0 || idx >= sheetPaths.size()) return;
        final int token = ++loadToken;
        final Uri uri = currentUri;
        final String path = sheetPaths.get(idx);
        final List<String> sst = shared;
        showMessage("Cargando hoja\u2026");
        new Thread(() -> {
            SheetData d = null;
            String err = null;
            try (ZipInputStream z = openEntry(uri, path)) {
                d = parseSheet(z, sst);
            } catch (Throwable e) { err = msg(e); }
            final SheetData fd = d;
            final String fe = err;
            ui.post(() -> {
                if (token != loadToken) return;
                if (fe != null || fd == null) { showMessage("No se pudo leer la hoja: " + fe); return; }
                rows = fd.rows;
                if (fd.truncated) Toast.makeText(this, "Se muestran las primeras " + MAX_ROWS + " filas", Toast.LENGTH_LONG).show();
                buildGrid();
                refresh = this::buildGrid;
            });
        }).start();
    }

    static int colIndex(String ref) {
        int idx = 0;
        for (int i = 0; i < ref.length(); i++) {
            char ch = ref.charAt(i);
            if (ch >= 'A' && ch <= 'Z') idx = idx * 26 + (ch - 'A' + 1);
            else if (ch >= 'a' && ch <= 'z') idx = idx * 26 + (ch - 'a' + 1);
            else break;
        }
        return idx - 1;
    }

    static String colName(int c) {
        StringBuilder s = new StringBuilder();
        c++;
        while (c > 0) { int m = (c - 1) % 26; s.insert(0, (char) ('A' + m)); c = (c - 1) / 26; }
        return s.toString();
    }

    private SheetData parseSheet(InputStream in, List<String> sst) throws Exception {
        SheetData d = new SheetData();
        XmlPullParser p = Xml.newPullParser();
        p.setInput(in, null);
        String[] row = null;
        int col = 0, lastCol = -1;
        String type = null;
        StringBuilder val = new StringBuilder();
        boolean inV = false, inT = false;
        for (int ev = p.getEventType(); ev != XmlPullParser.END_DOCUMENT; ev = p.next()) {
            if (ev == XmlPullParser.START_TAG) {
                String n = p.getName();
                if (n.equals("row")) {
                    String r = p.getAttributeValue(null, "r");
                    if (r != null) {
                        try {
                            int ri = Integer.parseInt(r);
                            while (d.rows.size() < ri - 1 && d.rows.size() < MAX_ROWS) d.rows.add(new String[MAX_COLS]);
                        } catch (NumberFormatException ignored) { }
                    }
                    if (d.rows.size() >= MAX_ROWS) { d.truncated = true; break; }
                    row = new String[MAX_COLS];
                    lastCol = -1;
                } else if (n.equals("c")) {
                    String ref = p.getAttributeValue(null, "r");
                    col = ref != null ? colIndex(ref) : lastCol + 1;
                    lastCol = col;
                    type = p.getAttributeValue(null, "t");
                    val.setLength(0);
                } else if (n.equals("v")) inV = true;
                else if (n.equals("t")) inT = true;
            } else if (ev == XmlPullParser.TEXT) {
                if (inV || inT) val.append(p.getText());
            } else if (ev == XmlPullParser.END_TAG) {
                String n = p.getName();
                if (n.equals("v")) inV = false;
                else if (n.equals("t")) inT = false;
                else if (n.equals("c")) {
                    if (row != null && col >= 0 && col < MAX_COLS) {
                        String s = cellText(type, val.toString(), sst);
                        if (s != null && s.length() > 0) row[col] = s;
                    }
                } else if (n.equals("row")) {
                    if (row != null) d.rows.add(row);
                    row = null;
                }
            }
        }
        return d;
    }

    private static String cellText(String type, String v, List<String> sst) {
        if (type == null) type = "n";
        switch (type) {
            case "s":
                try { return sst.get(Integer.parseInt(v.trim())); } catch (Exception e) { return ""; }
            case "b":
                return v.trim().equals("1") ? "VERDADERO" : "FALSO";
            case "str":
            case "inlineStr":
            case "e":
                return v;
            default:
                v = v.trim();
                if (v.isEmpty()) return "";
                try {
                    double dd = Double.parseDouble(v);
                    if (dd == Math.rint(dd) && Math.abs(dd) < 1e15) return String.valueOf((long) dd);
                    return new BigDecimal(v).round(new MathContext(10)).stripTrailingZeros().toPlainString();
                } catch (Exception e) { return v; }
        }
    }

    private static void cell(StringBuilder sb, String s, int width) {
        if (s == null) s = "";
        s = s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ');
        if (s.length() > width) s = s.substring(0, width - 1) + "\u2026";
        sb.append(s);
        for (int i = s.length(); i < width; i++) sb.append(' ');
    }

    private void buildGrid() {
        if (rows.isEmpty()) { showMessage("La hoja est\u00e1 vac\u00eda."); return; }
        int cols = 1;
        for (String[] r : rows) {
            for (int c = MAX_COLS - 1; c >= 0; c--) {
                if (r[c] != null) { cols = Math.max(cols, c + 1); break; }
            }
        }
        final int[] w = new int[cols];
        Arrays.fill(w, 3);
        for (String[] r : rows)
            for (int c = 0; c < cols; c++)
                if (r[c] != null) w[c] = Math.max(w[c], Math.min(MAX_COLW, r[c].length()));
        int total = 5;
        for (int c = 0; c < cols; c++) total += w[c] + 3;

        final float px = TypedValue.applyDimension(TypedValue.COMPLEX_UNIT_SP, textSp, getResources().getDisplayMetrics());
        Paint pt = new Paint();
        pt.setTypeface(Typeface.MONOSPACE);
        pt.setTextSize(px);
        final int widthPx = (int) (pt.measureText("0") * (total + 1)) + dp(16);
        final int fcols = cols;

        LinearLayout col = new LinearLayout(this);
        col.setOrientation(LinearLayout.VERTICAL);

        StringBuilder hb = new StringBuilder();
        cell(hb, "", 5);
        for (int c = 0; c < cols; c++) { hb.append(" \u2502 "); cell(hb, colName(c), w[c]); }
        TextView header = rowView(hb.toString(), 0xFFD0D0D0, true);
        col.addView(header, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.WRAP_CONTENT));

        ListView lv = new ListView(this);
        lv.setDividerHeight(0);
        lv.setAdapter(new BaseAdapter() {
            @Override public int getCount() { return rows.size(); }
            @Override public Object getItem(int p) { return null; }
            @Override public long getItemId(int p) { return p; }
            @Override public View getView(int pos, View cv, ViewGroup parent) {
                StringBuilder sb = new StringBuilder();
                cell(sb, String.valueOf(pos + 1), 5);
                String[] r = rows.get(pos);
                for (int c = 0; c < fcols; c++) { sb.append(" \u2502 "); cell(sb, r[c], w[c]); }
                TextView t = cv instanceof TextView ? (TextView) cv : rowView("", 0, false);
                t.setTextSize(TypedValue.COMPLEX_UNIT_SP, textSp);
                t.setBackgroundColor(pos % 2 == 0 ? 0xFFFFFFFF : 0xFFF2F2F2);
                t.setText(sb.toString());
                return t;
            }
        });
        col.addView(lv, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, 0, 1f));

        HorizontalScrollView hs = new HorizontalScrollView(this);
        hs.addView(col, new FrameLayout.LayoutParams(widthPx, ViewGroup.LayoutParams.MATCH_PARENT));
        content.removeAllViews();
        content.addView(hs, new FrameLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.MATCH_PARENT));
    }

    private TextView rowView(String text, int bg, boolean bold) {
        TextView t = new TextView(this);
        t.setTypeface(Typeface.MONOSPACE, bold ? Typeface.BOLD : Typeface.NORMAL);
        t.setTextSize(TypedValue.COMPLEX_UNIT_SP, textSp);
        t.setTextColor(0xFF111111);
        t.setSingleLine(true);
        t.setPadding(dp(8), dp(3), dp(8), dp(3));
        if (bg != 0) t.setBackgroundColor(bg);
        t.setText(text);
        return t;
    }
}
