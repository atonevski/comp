# encoding: UTF-8

module Comp

  CAT10 = %W{ 1f77b4 ff7f0e 2ca02c d62728 9467bd 8c564b e377c2 7f7f7f
              bcbd22 17becf }
  CAT20 = %W{ 1f77b4 aec7e8 ff7f0e ffbb78 2ca02c 98df8a d62728 ff9896
              9467bd c5b0d5 8c564b c49c94 e377c2 f7b6d2 7f7f7f c7c7c7
              bcbd22 dbdb8d 17becf 9edae5 }
  CAT20B = %W{ 393b79 5254a3 6b6ecf 9c9ede 637939 8ca252 b5cf6b cedb9c
               8c6d31 bd9e39 e7ba52 e7cb94 843c39 ad494a d6616b e7969c
               7b4173 a55194 ce6dbd de9ed6 }
  CAT20C = %W{ 3182bd 6baed6 9ecae1 c6dbef e6550d fd8d3c fdae6b fdd0a2
               31a354 74c476 a1d99b c7e9c0 756bb1 9e9ac8 bcbddc dadaeb
               636363 969696 bdbdbd d9d9d9 }

  YMD = '%Y-%m-%d'
  DMY = '%d.%m.%Y'

  def self.x_create_compare_sheet(book, day)
    rounded_period = {
      from: day - (day.day - 1),
      to:   day - (day.day % 7)
    }

    qry =<<-EOT
      g.id                    AS game_id,
      g.name                  AS name,
      CASE 
        WHEN g.type = 'INSTANT' THEN 1
        ELSE 0
      END                     AS is_instant,
      CASE
        WHEN g.parent IS NOT NULL THEN g.parent
        ELSE g.id
      END                     AS parent_id,
      SUM(s.sales)            AS sales,
      SUM(s.sales) / g.price  AS qty
    EOT
    sales = Sale.select(qry).
          joins('AS s INNER JOIN games AS g ON s.game_id = g.id').
          joins('INNER JOIN terminals AS t ON s.terminal_id = t.id')
    sales_a = sales.      
          where("s.date BETWEEN :from AND :to", rounded_period). 
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')
    sales_b = sales.      
          where("s.date BETWEEN :from AND :to", month_ago(rounded_period)).
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')
    sales_c = sales.      
          where("s.date BETWEEN :from AND :to", year_ago(rounded_period)).
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')
   
    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :center, vertical: :center },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_hdr_l = book.styles.add_style sz: 10,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_r = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 

    fmt_a = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :center }
    fmt_b = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left }
    fmt_c = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :right }
    fmt_d = fmt_c
    fmt_e = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#0.00%',
      alignment: { horizontal: :right }
    fmt_f = fmt_e

    fmt_ua = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :center }, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_ub = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left }, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_uc = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :right }, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_ud = fmt_uc 
    fmt_ue = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#0.00%',
      alignment: { horizontal: :right }, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_uf = fmt_ue

    fmt_tot = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }, b: true
    fmt_perc = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#0.00%', alignment: { horizontal: :right }, b: true

    book.add_worksheet name: 'Споредба по 7 денa' do |sheet|
      # table a
      sheet.add_row ["А: #{ rounded_period[:from].strftime DMY } --" +
                      " #{ rounded_period[:to].strftime DMY }",
                      '', '', '', '', ''], style: fmt_merge

      sheet.merge_cells 'A1:F1'

      sheet.add_row [ "id", "игра", "пари", "тикети / комб.", "А vs Б", "А vs В" ],
       style: [fmt_hdr_c, fmt_hdr_l, fmt_hdr_r, fmt_hdr_r, fmt_hdr_r, fmt_hdr_r],
       height: 24

      idx_inst = sales_a.find_index {|s| s.is_instant == 1}

      sales_a.each_with_index do |r, idx|
        style = (idx+1 == sales_a.length) ? # underline last ro
                [fmt_ua, fmt_ub, fmt_uc, fmt_ud, fmt_ue, fmt_uf] :
                [fmt_a, fmt_b, fmt_c, fmt_d, fmt_e, fmt_f]

        # a vs b
        b = sales_b.find_index {|s| r.game_id == s.game_id}
        a_vs_b = b ? "=(C" + (idx + 3).to_s + "-J" + (b + 3).to_s + ")/J" +
          (b + 3).to_s : nil
        # a vs c
        c = sales_c.find_index {|s| r.game_id == s.game_id}
        a_vs_c = c ? "=(C" + (idx + 3).to_s + "-O" + (c + 3).to_s + ")/O" +
          (c + 3).to_s : nil

        sheet.add_row [r.game_id, r.name, r.sales, r.qty, a_vs_b, a_vs_c],
          style: style, height: 12
      end
      # totals
      sheet.add_row [nil, 'лото', '=SUM(C3:C' + (idx_inst + 2).to_s + ')',
        '=SUM(D3:D' + (idx_inst + 2).to_s + ')',
        "=(C" + (sales_a.length + 3).to_s + "-J" + (sales_b.length + 3).to_s +
          ")/J" + (sales_b.length + 3).to_s,
        "=(C" + (sales_a.length + 3).to_s + "-O" + (sales_c.length + 3).to_s +
          ")/O" + (sales_c.length + 3).to_s],
          style: [fmt_tot]*4 + [fmt_perc]*2, height: 12
      sheet.add_row [nil, 'инстанти', '=SUM(C'+ (idx_inst + 3).to_s + ':C' +
        (sales_a.length + 2).to_s + ')',
        '=SUM(D'+ (idx_inst + 3).to_s + ':D' + (sales_a.length + 2).to_s + ')',
        "=(C" + (sales_a.length + 4).to_s + "-J" + (sales_b.length + 4).to_s +
          ")/J" + (sales_b.length + 4).to_s,
        "=(C" + (sales_a.length + 4).to_s + "-O" + (sales_c.length + 4).to_s +
          ")/O" + (sales_c.length + 4).to_s],
        style: [fmt_tot]*4 + [fmt_perc]*2, height: 12
      sheet.add_row [nil, 'вкупно',
        '=SUM(C3:C' + (sales_a.length + 2).to_s + ')', nil,
        "=(C" + (sales_a.length + 5).to_s + "-J" + (sales_b.length + 5).to_s +
          ")/J" + (sales_b.length + 5).to_s,
        "=(C" + (sales_a.length + 5).to_s + "-O" + (sales_c.length + 5).to_s +
          ")/O" + (sales_c.length + 5).to_s],
        style: [fmt_tot]*4 + [fmt_perc]*2, height: 12

      # append empty rows if needed
      ([sales_a.length, sales_b.length, sales_c.length].max - sales_a.length).times do
        sheet.add_row [nil]*6, height: 12
      end

      # b & c
      hdr     = [ "id", "игра", "пари", "тикети / комб." ]
      fmt_hdr = [ fmt_hdr_c, fmt_hdr_l, fmt_hdr_r, fmt_hdr_r ]
      fmt_row = [ fmt_a, fmt_b, fmt_c, fmt_d ]
      fmt_urow = [ fmt_ua, fmt_ub, fmt_uc, fmt_ud ]
      col     = %W{ game_id name sales qty }

      # table b
      a_month_ago = month_ago rounded_period
      sheet.rows[0].add_cell nil
      sheet.rows[0].add_cell "Б: #{ a_month_ago[:from].strftime DMY } --" +
                      " #{ a_month_ago[:to].strftime DMY }", style: fmt_merge
      sheet.rows[0].add_cell '', style: fmt_merge
      sheet.rows[0].add_cell '', style: fmt_merge
      sheet.rows[0].add_cell '', style: fmt_merge

      sheet.merge_cells 'H1:K1'

      sheet.rows[1].add_cell nil
      hdr.each_with_index do |h, i|
        sheet.rows[1].add_cell h, style: fmt_hdr[i]
      end

      idx_inst = sales_b.find_index {|s| s.is_instant == 1}

      sales_b.each_with_index do |r, idx|
        style = (idx+1 == sales_b.length ? fmt_urow : fmt_row) # underline last ro
        sheet.rows[idx + 2].add_cell nil
        (0..3).each do |i|
          sheet.rows[idx + 2].add_cell r[col[i]], style: style[i]
        end
      end

      # totals
      sheet.rows[sales_b.length + 2].add_cell nil
      sheet.rows[sales_b.length + 2].add_cell nil
      sheet.rows[sales_b.length + 2].add_cell 'лото', style: fmt_tot
      sheet.rows[sales_b.length + 2].add_cell '=SUM(J3:J' + (idx_inst + 2).to_s + ')',
        style: fmt_tot
      sheet.rows[sales_b.length + 2].add_cell '=SUM(K3:K' + (idx_inst + 2).to_s + ')',
        style: fmt_tot

      sheet.rows[sales_b.length + 3].add_cell nil
      sheet.rows[sales_b.length + 3].add_cell nil
      sheet.rows[sales_b.length + 3].add_cell 'инстанти', style: fmt_tot
      sheet.rows[sales_b.length + 3].add_cell '=SUM(J'+ (idx_inst + 3).to_s + ':J' +
        (sales_b.length + 2).to_s + ')', style: fmt_tot
      sheet.rows[sales_b.length + 3].add_cell '=SUM(K'+ (idx_inst + 3).to_s + ':K' +
        (sales_b.length + 2).to_s + ')', style: fmt_tot

      sheet.rows[sales_b.length + 4].add_cell nil
      sheet.rows[sales_b.length + 4].add_cell nil
      sheet.rows[sales_b.length + 4].add_cell 'вкупно', style: fmt_tot
      sheet.rows[sales_b.length + 4].add_cell '=SUM(J3:J' +
        (sales_b.length + 2).to_s + ')', style: fmt_tot
      sheet.rows[sales_b.length + 4].add_cell nil

      # empty rows if needed
      (0 .. sales_c.length - sales_b.length - 1).each do |i|
        sheet.rows[sales_b.length + 5 + i].add_cell nil
        sheet.rows[sales_b.length + 5 + i].add_cell nil
        sheet.rows[sales_b.length + 5 + i].add_cell nil
        sheet.rows[sales_b.length + 5 + i].add_cell nil
        sheet.rows[sales_b.length + 5 + i].add_cell nil
      end

      # table c
      a_year_ago = year_ago rounded_period
      sheet.rows[0].add_cell nil
      sheet.rows[0].add_cell "В: #{ a_year_ago[:from].strftime DMY } --" +
                      " #{ a_year_ago[:to].strftime DMY }", style: fmt_merge
      sheet.rows[0].add_cell '', style: fmt_merge
      sheet.rows[0].add_cell '', style: fmt_merge
      sheet.rows[0].add_cell '', style: fmt_merge

      sheet.merge_cells 'M1:P1'

      sheet.rows[1].add_cell nil
      hdr.each_with_index do |h, i|
        sheet.rows[1].add_cell h, style: fmt_hdr[i]
      end

      idx_inst = sales_c.find_index {|s| s.is_instant == 1}

      sales_c.each_with_index do |r, idx|
        style = (idx+1 == sales_c.length ? fmt_urow : fmt_row) # underline last ro
        sheet.rows[idx + 2].add_cell nil
        (0..3).each do |i|
          sheet.rows[idx + 2].add_cell r[col[i]], style: style[i]
        end
      end

      # totals
      sheet.rows[sales_c.length + 2].add_cell nil
      sheet.rows[sales_c.length + 2].add_cell nil
      sheet.rows[sales_c.length + 2].add_cell 'лото', style: fmt_tot
      sheet.rows[sales_c.length + 2].add_cell '=SUM(O3:O' + (idx_inst + 2).to_s + ')',
        style: fmt_tot
      sheet.rows[sales_c.length + 2].add_cell '=SUM(P3:P' + (idx_inst + 2).to_s + ')',
        style: fmt_tot
      sheet.rows[sales_c.length + 2].add_cell nil
      sheet.rows[sales_c.length + 2].add_cell nil

      sheet.rows[sales_c.length + 3].add_cell nil
      sheet.rows[sales_c.length + 3].add_cell nil
      sheet.rows[sales_c.length + 3].add_cell 'инстанти', style: fmt_tot
      sheet.rows[sales_c.length + 3].add_cell '=SUM(O'+ (idx_inst + 3).to_s + ':O' +
        (sales_c.length + 2).to_s + ')', style: fmt_tot
      sheet.rows[sales_c.length + 3].add_cell '=SUM(P'+ (idx_inst + 3).to_s + ':P' +
        (sales_c.length + 2).to_s + ')', style: fmt_tot
      sheet.rows[sales_c.length + 3].add_cell nil
      sheet.rows[sales_c.length + 3].add_cell nil

      sheet.rows[sales_c.length + 4].add_cell nil
      sheet.rows[sales_c.length + 4].add_cell nil
      sheet.rows[sales_c.length + 4].add_cell 'вкупно', style: fmt_tot
      sheet.rows[sales_c.length + 4].add_cell '=SUM(O3:O' +
        (sales_c.length + 2).to_s + ')', style: fmt_tot
      sheet.rows[sales_c.length + 4].add_cell nil
      sheet.rows[sales_c.length + 4].add_cell nil
      sheet.rows[sales_c.length + 4].add_cell nil

      # fix column widths
      sheet.column_widths 3.8, 16, 12, 10, 8, 8, 6,
        3.8, 16, 12, 10, 6,
        3.8, 16, 12, 10, 6

      # add the chart
      sheet.add_row
      sheet.add_row [nil, nil, 'А', 'Б', 'В']
      sheet.add_row [nil, 'лото', '=C' + (sales_a.length + 3).to_s,
        '=J' + (sales_b.length + 3).to_s,
        '=O' + (sales_c.length + 3).to_s], style: fmt_c
      sheet.add_row [nil, 'инстанти', '=C' + (sales_a.length + 4).to_s,
        '=J' + (sales_b.length + 4).to_s,
        '=O' + (sales_c.length + 4).to_s], style: fmt_c

      r = [sales_a.length, sales_b.length, sales_c.length].max + 7
      h = 15
      sheet.add_chart Axlsx::Bar3DChart, bg_color: '88FF88',
        start_at: 'A' + r.to_s, end_at: 'I' + (r + h - 1).to_s do |chart|
        chart.title     = 'Споредба по 7 денa'
        chart.grouping = :stacked
        chart.add_series data: sheet["C#{r + 1}:E#{r + 1}"],
          labels: sheet["C#{r}:E#{r}"], title: sheet["B#{r + 1}"],
          colors: [CAT10[0]]*3, color: '000000'
        chart.add_series data: sheet["C#{r + 2}:E#{r + 2}"],
          labels: sheet["C#{r}:E#{r}"], title: sheet["B#{r + 2}"],
          colors: [CAT10[1]]*3, color: '000000'
        # chart.valAxis.title = 'Y Axis' 
        chart.show_legend = true
        chart.bar_dir = :col
        chart.valAxis.gridlines = true
        chart.catAxis.gridlines = true
        chart.val_axis.format_code = "#,###"
      end

    end
  end # compare sheet

  def self.x_create_share_sheet(book, day)
    qry =<<-EOT
      g.id                    AS game_id,
      g.name                  AS name,
      g.price                 AS price,
      CASE 
        WHEN g.type = 'INSTANT' THEN 1
        ELSE 0
      END                     AS is_instant,
      CASE
        WHEN g.parent IS NOT NULL THEN g.parent
        ELSE g.id
      END                     AS parent_id,
      SUM(s.sales)            AS sales,
      SUM(s.sales) / g.price  AS qty,
      COUNT(DISTINCT t.id)    AS term_count
    EOT

    sales = Sale.select(qry).
          joins('AS s INNER JOIN games AS g ON s.game_id = g.id').
          joins('INNER JOIN terminals AS t ON s.terminal_id = t.id')
    sales_i = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day). 
          where('t.agent_id <> 225').
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')
    sales_d = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day). 
          where('t.agent_id = 225').
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')

    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :center, vertical: :center },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_hdr_l = book.styles.add_style sz: 10,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_r = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 

    fmt_a = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :center }
    fmt_b = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left }
    fmt_c = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :right }
    fmt_d = fmt_c
    fmt_e = fmt_c
    fmt_f = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#0.00%',
      alignment: { horizontal: :right }

    fmt_ua = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :center },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_ub = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_uc = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :right },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_ud = fmt_uc
    fmt_ue = fmt_uc
    fmt_uf = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#0.00%',
      alignment: { horizontal: :right },
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_tot = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }, b: true
    fmt_perc = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#0.00%', alignment: { horizontal: :right }, b: true

    fmt_row   = [fmt_a, fmt_b, fmt_c, fmt_d, fmt_e, fmt_f]
    fmt_urow  = [fmt_ua, fmt_ub, fmt_uc, fmt_ud, fmt_ue, fmt_uf]

    book.add_worksheet name: 'Учество' do |sheet|
      # ind & dir headers
      sheet.add_row [
        "Индиректна продажба, #{ MONTH_NAMES_MK[day.month] }" +
          " #{ day.year }", '', '', '', '', '', '',
        "Директна продажба, " +
        "#{ MONTH_NAMES_MK[day.month] } #{ day.year }", '', '', '', '', ''],
        style: [fmt_merge]*6 + [nil] + [fmt_merge]*6
      sheet.merge_cells 'A1:F1'
      sheet.merge_cells 'H1:M1'
      sheet.add_row [ "id", "игра", "пари", "тикети / комб.", "бр. терм.",
        "удел", '', "id", "игра", "пари", "тикети / комб.", "бр. терм.", "удел"],
        style: [fmt_hdr_c, fmt_hdr_l, fmt_hdr_r, fmt_hdr_r, fmt_hdr_r, fmt_hdr_r, nil,
          fmt_hdr_c, fmt_hdr_l, fmt_hdr_r, fmt_hdr_r, fmt_hdr_r, fmt_hdr_r],
        height: 24

      # ind
      idx_inst = sales_i.find_index {|s| s.is_instant == 1}
      sales_i.each_with_index do |s, idx|
        style = (idx+1 == sales_i.length) ? fmt_urow : fmt_row # underline last ro

        # ind vs dir 
        d = sales_d.find_index {|r| r.game_id == s.game_id}
        share = d ? "=C" + (idx + 3).to_s + "/(J" + (d + 3).to_s +
          "+C" + (idx + 3).to_s + ")" : nil
        sheet.add_row [s.game_id, s.name, s.sales, s.qty, s.term_count, share],
          style: style, height: 12
      end 

      # totals
      sheet.add_row [nil, 'лото', '=SUM(C3:C' + (idx_inst + 2).to_s + ')',
        '=SUM(D3:D' + (idx_inst + 2).to_s + ')', nil, "=C" +
          (sales_i.length + 3).to_s + "/(C" + (sales_i.length + 3).to_s +
          " + J" + (sales_d.length + 3).to_s + ")"],
        style: [fmt_tot]*5 + [fmt_perc]*1, height: 12
      sheet.add_row [nil, 'инстанти', '=SUM(C'+ (idx_inst + 3).to_s + ':C' +
        (sales_i.length + 2).to_s + ')',
        '=SUM(D'+ (idx_inst + 3).to_s + ':D' + (sales_i.length + 2).to_s + ')',
        nil, "=C" + (sales_i.length + 4).to_s + "/(C" +
        (sales_i.length + 4).to_s + " + J" + (sales_d.length + 4).to_s + ")"],
        style: [fmt_tot]*5 + [fmt_perc]*1, height: 12
      sheet.add_row [nil, 'вкупно',
        '=SUM(C3:C' + (sales_i.length + 2).to_s + ')', nil, nil,
        "=C" + (sales_i.length + 5).to_s + "/(C" + (sales_i.length + 5).to_s +
          " + J" + (sales_d.length + 5).to_s + ")"],
        style: [fmt_tot]*4 + [fmt_perc]*2, height: 12

      # append empty rows if needed
      ([sales_i.length, sales_d.length].max - sales_i.length).times do
        sheet.add_row [nil]*6, height: 12
      end
      
      sheet.add_row []
      sheet.add_row [nil, 'индиректна лото', '=C' + (sales_i.length + 3).to_s]
      sheet.add_row [nil, 'индиректна инстанти', '=C' + (sales_i.length + 4).to_s]
      sheet.add_row [nil, 'директна лото', '=J' + (sales_i.length + 3).to_s]
      sheet.add_row [nil, 'директна инстанти', '=J' + (sales_i.length + 4).to_s]
      
      r = [sales_i.length, sales_d.length].max + 6 + 1
      h = 15
      sheet.add_chart Axlsx::Pie3DChart, start_at: "A#{r}",
        end_at: "G#{r+h}" do |chart|
       
        chart.title = 'Учество во вкупен промет'

        ser = chart.add_series data: sheet["C#{r}:C#{r+3}"],
           labels: sheet["B#{r}:B#{r+3}"], 
           colors: CAT10

        chart.d_lbls.show_percent = true
        chart.show_legend = true
        chart.d_lbls.d_lbl_pos = :outEnd
      end

      # dir
      idx_inst  = sales_d.find_index {|s| s.is_instant == 1}
      col       = %W{ game_id name sales qty term_count }
      sales_d.each_with_index do |s, idx|
        style = (idx+1 == sales_d.length) ? fmt_urow : fmt_row # underline last ro

        sheet.rows[idx + 2].add_cell nil
        col.length.times do |c|
          sheet.rows[idx + 2].add_cell s[col[c]], style: style[c]
        end
        # ind vs dir 
        i = sales_i.find_index {|r| r.game_id == s.game_id}
        share = i ? "=J" + (idx + 3).to_s + "/(J" + (i + 3).to_s +
           "+C" + (idx + 3).to_s + ")" : nil
        sheet.rows[idx + 2].add_cell share, style: style[5]
      end

      # totals
      sheet.rows[sales_d.length + 2].add_cell nil
      sheet.rows[sales_d.length + 2].add_cell nil
      sheet.rows[sales_d.length + 2].add_cell 'лото', style: fmt_tot
      sheet.rows[sales_d.length + 2].add_cell '=SUM(J3:J' + (idx_inst + 2).to_s + ')',
        style: fmt_tot
      sheet.rows[sales_d.length + 2].add_cell '=SUM(K3:K' + (idx_inst + 2).to_s + ')',
        style: fmt_tot
      sheet.rows[sales_d.length + 2].add_cell nil
      sheet.rows[sales_d.length + 2].add_cell "=J" + (sales_d.length + 3).to_s +
        "/(C" + (sales_i.length + 3).to_s + " + J" + (sales_d.length + 3).to_s + ")",
        style: fmt_perc

      
      sheet.rows[sales_d.length + 3].add_cell nil
      sheet.rows[sales_d.length + 3].add_cell nil
      sheet.rows[sales_d.length + 3].add_cell 'инстанти', style: fmt_tot
      sheet.rows[sales_d.length + 3].add_cell '=SUM(J'+ (idx_inst + 3).to_s + ':J' +
        (sales_d.length + 2).to_s + ')', style: fmt_tot
      sheet.rows[sales_d.length + 3].add_cell '=SUM(K'+ (idx_inst + 3).to_s + ':K' +
        (sales_d.length + 2).to_s + ')', style: fmt_tot
      sheet.rows[sales_d.length + 3].add_cell nil
      sheet.rows[sales_d.length + 3].add_cell "=J" + (sales_d.length + 4).to_s +
        "/(C" + (sales_i.length + 4).to_s + " + J" + (sales_d.length + 4).to_s + ")",
        style: fmt_perc


      sheet.rows[sales_d.length + 4].add_cell nil
      sheet.rows[sales_d.length + 4].add_cell nil
      sheet.rows[sales_d.length + 4].add_cell 'вкупно', style: fmt_tot
      sheet.rows[sales_d.length + 4].add_cell '=SUM(J3:J' +
        (sales_d.length + 2).to_s + ')', style: fmt_tot
      sheet.rows[sales_d.length + 4].add_cell nil
      sheet.rows[sales_d.length + 4].add_cell nil
      sheet.rows[sales_d.length + 4].add_cell "=J" + (sales_d.length + 5).to_s +
        "/(C" + (sales_i.length + 5).to_s + " + J" + (sales_d.length + 5).to_s + ")",
        style: fmt_perc

      # fix column widths
      sheet.column_widths 3.8, 16, 12, 10, 8, 8, 6,
        3.8, 16, 12, 10, 8, 8, 6
    end # add_worksheet
  end # share sheet

  def self.x_create_remainder_sheet(book, day)
    funds_file      = File.expand_path '../../../config/funds.yml', __FILE__
    commission_file = File.expand_path '../../../config/commission.yml', __FILE__

    funds       = YAML::load File.open funds_file
    commission  = YAML::load File.open commission_file

    qry =<<-EOT
      g.id                    AS game_id,
      g.name                  AS name,
      g.price                 AS price,
      CASE 
        WHEN g.type = 'INSTANT' THEN 1
        ELSE 0
      END                     AS is_instant,
      CASE
        WHEN g.parent IS NOT NULL THEN g.parent
        ELSE g.id
      END                     AS parent_id,
      SUM(s.sales)            AS sales,
      SUM(s.sales) / g.price  AS qty,
      COUNT(DISTINCT t.id)    AS term_count
    EOT

    sales = Sale.select(qry).
          joins('AS s INNER JOIN games AS g ON s.game_id = g.id').
          joins('INNER JOIN terminals AS t ON s.terminal_id = t.id')
    sales_i = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day). 
          where('t.agent_id <> 225').
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')
    sales_d = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day). 
          where('t.agent_id = 225').
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')

    # formating
    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :center, vertical: :center },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_hdr_l = book.styles.add_style sz: 10,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_r = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 


    fmt_a = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :center }
    fmt_b = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left }
    fmt_c = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :center }
    fmt_d = fmt_c
    fmt_e = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#0.00%',
      alignment: { horizontal: :right }
    fmt_f = fmt_e
    fmt_g = book.styles.add_style sz: 10, font_name: 'Droid Sans', format_code: '#,###',
      alignment: { horizontal: :right }
    fmt_h = fmt_g
    fmt_i = fmt_g
    fmt_j = fmt_g

    fmt_ua = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :center },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_ub = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_uc = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :center },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_ud = fmt_uc
    fmt_ue = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#0.00%', alignment: { horizontal: :right },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_uf = fmt_ue
    fmt_ug = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_uh = fmt_ug
    fmt_ui = fmt_ug
    fmt_uj = fmt_ug

    fmt_irow = [fmt_a, fmt_b, fmt_c, fmt_d, fmt_e, fmt_f, fmt_g,
      fmt_h, fmt_i, fmt_j] 
    fmt_iurow = [fmt_ua, fmt_ub, fmt_uc, fmt_ud, fmt_ue, fmt_uf, fmt_ug,
      fmt_uh, fmt_ui, fmt_uj] 

    fmt_drow = [fmt_a, fmt_b, fmt_c, fmt_d, fmt_e, fmt_g, fmt_h, fmt_i, fmt_j] 
    fmt_durow = [fmt_ua, fmt_ub, fmt_uc, fmt_ud, fmt_ue, fmt_ug,
      fmt_uh, fmt_ui, fmt_uj] 

    fmt_tot = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }, b: true
    fmt_perc = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#0.00%', alignment: { horizontal: :right }, b: true

    # render worksheet
    book.add_worksheet name: 'Остаток' do |sheet|
      # ind & dir headers
      sheet.add_row [
        "Индиректна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }",
        *['']*9, nil,
        "Директна продажба, " +
        "#{ MONTH_NAMES_MK[day.month] } #{ day.year }", *['']*8],
        style: [fmt_merge]*10 + [nil] + [fmt_merge]*9
      sheet.merge_cells 'A1:J1'
      sheet.merge_cells 'L1:T1'
      sheet.add_row [ "id", "игра", "цена", "терм.", "фонд за доб.", "пров.",
        "уплата", "тикети / комб.", "фонд + пров. + МПМ + РМ", "остаток", nil,
        "id", "игра", "цена", "терм.", "фонд за доб.", "уплата",
        "тикети / комб.", "фонд + РМ", "остаток", ],
        style: [fmt_hdr_c, fmt_hdr_l] + [fmt_hdr_c]*8 + [nil, fmt_hdr_c,
          fmt_hdr_l] + [fmt_hdr_c]*7,
        height: 34

      # ind table alone
      sales_i.each_with_index do |s, idx|
        unless funds[s.game_id.to_s]
          raise "Undefined funds for #{ s.game_id }"
        end
        unless commission[s.game_id.to_s]
          raise "Undefined commission for #{ s.game_id }"
        end
        style = (idx+1 == sales_i.length) ? fmt_iurow : fmt_irow # underline last ro
        sheet.add_row [s.game_id, s.name, s.price, s.term_count,
          funds[s.game_id.to_s], commission[s.game_id.to_s], s.sales, s.qty,
          "=(E#{ idx + 3 } + F#{ idx + 3 } + #{ RM_PERC } + #{ MPM_PERC })*G#{ idx + 3 }",
          "=G#{ idx + 3 } - I#{ idx + 3 }"],
          style: style, height: 12
      end
      sheet.add_row [nil]*6 + ["=SUM(G3:G#{ sales_i.length + 2 })", nil,
        "=SUM(I3:I#{ sales_i.length + 2 })", 
        "=SUM(J3:J#{ sales_i.length + 2 })"], style: fmt_tot, height: 12
      sheet.add_row [nil]*9 + 
        ["=J#{ sales_i.length + 3 }/G#{ sales_i.length + 3}"],
        style: fmt_perc, height: 12
      
      # add empty rows if needed, i.e. sales_i & sales_d sizes differ
      ([sales_i.length, sales_d.length].max - sales_i.length).times do 
        sheet.add_row nil
      end

      # dir table alone
      cols = %W[ game_id name price term_count ]

      sales_d.each_with_index do |s, idx|
        unless funds[s.game_id.to_s]
          raise "Undefined funds for #{ s.game_id }"
        end
        style = (idx+1 == sales_d.length) ? fmt_durow : fmt_drow # underline last ro
        
        sheet.rows[idx + 2].add_cell nil
        cols.length.times do |c|
          sheet.rows[idx + 2].add_cell s[cols[c]], style: style[c]
        end
        sheet.rows[idx + 2].add_cell funds[s.game_id.to_s], style: style[4]
        sheet.rows[idx + 2].add_cell s.sales, style: style[5]
        sheet.rows[idx + 2].add_cell s.qty, style: style[6]
        sheet.rows[idx + 2].add_cell "=(P#{ idx + 3 } + #{ RM_PERC })*Q#{ idx + 3 }",
          style: style[7]
        sheet.rows[idx + 2].add_cell "=Q#{ idx + 3 } - S#{ idx + 3 }",
          style: style[8]
      end
      # dir totals
      6.times { sheet.rows[sales_d.length + 2].add_cell nil }
      sheet.rows[sales_d.length + 2].add_cell "=SUM(Q3:Q#{sales_d.length + 2})",
        style: fmt_tot
      sheet.rows[sales_d.length + 2].add_cell nil 
      sheet.rows[sales_d.length + 2].add_cell "=SUM(S3:S#{sales_d.length + 2})",
        style: fmt_tot
      sheet.rows[sales_d.length + 2].add_cell "=SUM(T3:T#{sales_d.length + 2})",
        style: fmt_tot

      9.times { sheet.rows[sales_d.length + 3].add_cell nil }
      sheet.rows[sales_d.length + 3].add_cell( 
        "=T#{sales_d.length + 3}/Q#{sales_d.length + 3}", style: fmt_perc)

      # charts
      sheet.add_row 

      # chart i
      sheet.add_row [nil, 'фонд',
        sales_i.inject(0){|sum, r| sum + r.sales * funds[r.game_id.to_s]}] +
          [nil]*8, style: fmt_g
      sheet.add_row [nil, 'провизија',
        sales_i.inject(0){|sum, r| sum + r.sales * commission[r.game_id.to_s]}]+
        [nil]*8, style: fmt_g
      sheet.add_row [nil, 'МПМ', "=G#{ sales_i.length + 3 } * #{ MPM_PERC }"] +
        [nil]*8, style: fmt_g
      sheet.add_row [nil, 'РМ', "=G#{ sales_i.length + 3 } * #{ RM_PERC }"] +
        [nil]*8, style: fmt_g
      sheet.add_row [nil, 'остаток', "=J#{ sales_i.length + 3 }"] +
        [nil]*8, style: fmt_g

      r = [sales_i.length, sales_d.length].max + 6
      h = 15
      sheet.add_chart Axlsx::Pie3DChart, start_at: "A#{r}",
        end_at: "I#{r + h}" do |chart|
       
        chart.title = 'Структура на индиректна продажба'

        ser = chart.add_series data: sheet["C#{r}:C#{r+4}"],
           labels: sheet["B#{r}:B#{r+4}"],
           colors: CAT10

        chart.d_lbls.show_percent = true
        chart.show_legend = true
        chart.d_lbls.show_leader_lines = true
        chart.d_lbls.d_lbl_pos = :outEnd
      end

      # chart d
      sheet.rows[r-1].add_cell nil
      sheet.rows[r-1].add_cell 'фонд', style: fmt_g
      sheet.rows[r-1].add_cell sales_d.inject(0){|sum, r|
        sum + r.sales * funds[r.game_id.to_s]}, style: fmt_g

      sheet.rows[r].add_cell nil
      sheet.rows[r].add_cell 'РМ', style: fmt_g
      sheet.rows[r].add_cell "=Q#{ sales_d.length + 3 } * #{ RM_PERC }",
        style: fmt_g

      sheet.rows[r+1].add_cell nil
      sheet.rows[r+1].add_cell 'остаток', style: fmt_g
      sheet.rows[r+1].add_cell "=T#{sales_d.length + 3}",
        style: fmt_g

      sheet.add_chart Axlsx::Pie3DChart, start_at: "L#{r}",
        end_at: "T#{r + h}" do |chart|
       
        chart.title = 'Структура на директна продажба'

        ser = chart.add_series data: sheet["N#{r}:N#{r+2}"],
           labels: sheet["M#{r}:M#{r+2}"],
           colors: CAT10

        chart.d_lbls.show_percent = true
        chart.show_legend = true
        chart.d_lbls.show_leader_lines = true
        chart.d_lbls.d_lbl_pos = :outEnd
      end

      # fix column widths
      sheet.column_widths 3.8, 16, 4.5, 5, 8, 6, 12, 10, 12, 12, 6,
        3.8, 16, 4.5, 5, 8, 10, 10, 10, 12
    end
  end # remainder sheet

  # Weekly sales
  #
  def self.x_create_weekly_sheet(book, day)
    qry =<<-EOT
      date(s.date, 'weekday 0', '-6 days')      AS monday,
      date(s.date, 'weekday 0')                 AS sunday,
      CAST(strftime('%W', date(s.date, 'weekday 0')) AS INTEGER)
                                                AS week_number,
      SUM(s.sales)                              AS sales
      --- FROM sales AS s
    EOT

    wsales_now = Sale.select(qry).
          joins('AS s').
          where('substr(sunday, 1, 4) = :year_s', year_s: day.year.to_s).
          group('monday', 'sunday').having('sunday <= :day', day: day).
          order('monday')
    
    last_week = wsales_now[-1].week_number 

    wsales_before = Sale.select(qry).
          joins('AS s').
          where('substr(sunday, 1, 4) = :year_s', year_s: (day.year - 1).to_s).
          group('monday', 'sunday').
          having('week_number <= :week_number', week_number: last_week).
          order('monday')

    wsales_before2 = Sale.select(qry).
          joins('AS s').
          where('substr(sunday, 1, 4) = :year_s', year_s: (day.year - 2).to_s).
          group('monday', 'sunday').
          having('week_number <= :week_number', week_number: last_week).
          order('monday')

    # formating
    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :center, vertical: :center },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 

    fmt_a = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '00', alignment: { horizontal: :center }
    fmt_b = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: 'DD.MM.YYYY', alignment: { horizontal: :right }
    fmt_c = fmt_b
    fmt_d = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }

    # render worksheet
    book.add_worksheet name: 'Неделно' do |sheet|
      # cur & year before headers
      sheet.add_row [
        "#{ day.year }", *['']*3, nil,
        "#{ day.year - 1 }", *['']*3, nil,
        "#{ day.year - 2 }", *['']*3, nil, ],
        style: [fmt_merge]*4 + [nil] +
               [fmt_merge]*4 + [nil] +
               [fmt_merge]*4 + [nil]
      sheet.merge_cells 'A1:D1'
      sheet.merge_cells 'F1:I1'
      sheet.merge_cells 'K1:N1'

      sheet.add_row [ "седмица", "понед.", "недела", "уплата", "",
                      "седмица", "понед.", "недела", "уплата", "",
                      "седмица", "понед.", "недела", "уплата", ],
        style: [fmt_hdr_c]*4 + [nil] + [fmt_hdr_c]*4 +
               [nil] + [fmt_hdr_c]*4 #, height: 34

      # fix column widths
      sheet.column_widths *[10]*4, 6, *[10]*4, 6, *[10]*4, 6

      wsales_now.each_with_index do |s, idx|
        b  = wsales_before[idx]
        b2 = wsales_before2[idx]
        sheet.add_row [
          s.week_number, Date.parse(s.monday), 
          Date.parse(s.sunday), s.sales, nil,

          b.week_number, Date.parse(b.monday), 
          Date.parse(b.sunday), b.sales, nil,

          b2.week_number, Date.parse(b2.monday), 
          Date.parse(b2.sunday), b2.sales, nil],
          style: [fmt_a, fmt_b, fmt_c, fmt_d, nil,
           fmt_a, fmt_b, fmt_c, fmt_d, nil,
           fmt_a, fmt_b, fmt_c, fmt_d, nil],
          height: 12
      end

      # render chart
      sheet.add_chart Axlsx::LineChart, 
        start_at: 'P1', end_at: 'Y24' do |chart|
         
        chart.title = "Неделен промет #{ day.year }/#{ day.year - 1}/#{ day.year - 2 }"
        chart.add_series data: sheet["D3:D#{ wsales_now.length + 2 }"],
          labels: sheet["A3:A#{ wsales_now.length + 2 }"],
          color: CAT10[0], title: sheet['A1']
        chart.add_series data: sheet["I3:I#{ wsales_now.length + 2 }"],
          # labels: sheet["A3:A#{ wsales_now.length + 2 }"],
          color: CAT10[1], title: sheet['F1']
        chart.add_series data: sheet["N3:N#{ wsales_now.length + 2 }"],
          # labels: sheet["A3:A#{ wsales_now.length + 2 }"],
          color: CAT10[2], title: sheet['K1']
        chart.valAxis.gridlines = false
        chart.catAxis.gridlines = false
        chart.val_axis.format_code = "#,###"
        chart.cat_axis.tick_lbl_skip = 10
        # chart.cat_axis.tick_mark_skip = 9
        chart.cat_axis.color = '000000'
        chart.val_axis.color = '000000'
        chart.cat_axis.tick_lbl_pos = :nextTo
        chart.cat_axis.crosses = :min
        chart.cat_axis.auto = false
      end
    end
  end # weekly sales

  ##
  # Compare months sheet
  def Comp.x_create_compare_months_sheet book, day
    # 1st we get the sales for A: current month, B: previous month, C: a year ago
    qry =<<-EOT
      g.id                    AS game_id,
      g.name                  AS name,
      g.price                 AS price,
      CASE 
        WHEN g.type = 'INSTANT' THEN 1
        ELSE 0
      END                     AS is_instant,
      CASE
        WHEN g.parent IS NOT NULL THEN g.parent
        ELSE g.id
      END                     AS parent_id,
      SUM(s.sales)            AS sales,
      SUM(s.sales) / g.price  AS qty,
      COUNT(DISTINCT t.id)    AS term_count
    EOT
    
    # use same pattern for date interval
    day_b = if day.month == 1 # month before
              Date.new day.year - 1, 12, 1
            else
              Date.new day.year, day.month - 1, 1
            end
    day_c = Date.new day.year - 1, day.month, 1 # year before (same month)

    sales = Sale.select(qry).
          joins('AS s INNER JOIN games AS g ON s.game_id = g.id').
          joins('INNER JOIN terminals AS t ON s.terminal_id = t.id')
    # A:
    sales_a = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND" +
            " date(:day,'start of month','+1 month','-1 day')", day: day). 
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')
    # B:
    sales_b = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND" +
            " date(:day,'start of month','+1 month','-1 day')", day: day_b). 
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')
    # C:
    sales_c = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND" +
            " date(:day,'start of month','+1 month','-1 day')", day: day_c). 
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')
    cols = %W{ game_id name term_count sales qty }

    # formating
    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :center, vertical: :center },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_hdr_l = book.styles.add_style sz: 10,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_r = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr = [fmt_hdr_c, fmt_hdr_l, fmt_hdr_c, fmt_hdr_r, fmt_hdr_r, nil]

    fmt_a = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :center }
    fmt_b = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left }
    fmt_c = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :center }
    fmt_d = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }
    fmt_e = fmt_d

    fmt_ua = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :center },
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_ub = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_uc = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :center },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_ud = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_ue = fmt_ud

    fmt_row  = [fmt_a, fmt_b, fmt_c, fmt_d, fmt_e, nil]
    fmt_urow   = [fmt_ua, fmt_ub, fmt_uc, fmt_ud, fmt_ue, nil]

    fmt_tot = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }, b: true
    fmt_perc = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#0.00%', alignment: { horizontal: :right }, b: true

    # render worksheet
    book.add_worksheet name: 'Споредба на месеци' do |sheet|
      sheet.add_row ["#{ MONTH_NAMES_MK[day.month] } #{ day.year }"] +
        ['']*4 + [nil] +
        ["#{ MONTH_NAMES_MK[day_b.month] } #{ day_b.year }"] +
        ['']*4 + [nil] +
        ["#{ MONTH_NAMES_MK[day_c.month] } #{ day_c.year }"] +
        ['']*4 + [nil],
        style: ([fmt_merge]*5 + [nil])*3 

      sheet.merge_cells 'A1:E1'
      sheet.merge_cells 'G1:K1'
      sheet.merge_cells 'M1:Q1'

      sheet.add_row [ "id", "игра", "бр.\nтерм.", "уплата", "комб./ тикети", '']*3, 
        style: fmt_hdr*3, height: 24

      # A 
      idx_inst = sales_a.find_index {|s| s.is_instant == 1}
      sales_a.each_with_index do |s, idx|
        style = (idx+1 == sales_a.length) ? fmt_urow : fmt_row 
        sheet.add_row [s.game_id, s.name, s.term_count, s.sales, s.qty, nil],
          style: style, height: 12
      end
      # totals
      sheet.add_row [nil, nil, 'лото', "=SUM(D3:D#{idx_inst + 2})",
        "=SUM(E3:E#{idx_inst + 2})", nil],
        style: [fmt_tot]*5 + [nil], height: 12
      sheet.add_row [nil, nil, 'инстанти', 
        "=SUM(D#{idx_inst + 3}:D#{sales_a.length + 2})",
        "=SUM(E#{idx_inst + 3}:E#{sales_a.length + 2})", nil],
        style: [fmt_tot]*5 + [nil], height: 12
      sheet.add_row [nil, nil, 'вкупно',
        "=SUM(D3:D#{sales_a.length + 2})", nil, nil],
        style: [fmt_tot]*5 + [nil], height: 12
      # append empty rows if needed
      ([sales_a.length, sales_b.length, sales_c.length].max - sales_a.length).times do
        sheet.add_row [nil]*6, height: 12
      end

      # B 
      idx_inst = sales_b.find_index {|s| s.is_instant == 1}
      sales_b.each_with_index do |s, idx|
        style = (idx+1 == sales_b.length) ? fmt_urow : fmt_row
        cols.each_with_index do |c, i|
          sheet.rows[idx + 2].add_cell s[c], style: style[i]
        end
        sheet.rows[idx + 2].add_cell nil
      end
      # totals
      2.times { sheet.rows[sales_b.length+2].add_cell nil }
      sheet.rows[sales_b.length+2].add_cell 'лото', style: fmt_tot
      sheet.rows[sales_b.length+2].add_cell "=SUM(J3:J#{idx_inst + 2})",
        style: fmt_tot
      sheet.rows[sales_b.length+2].add_cell "=SUM(K3:K#{idx_inst + 2})",
        style: fmt_tot
      sheet.rows[sales_b.length+2].add_cell nil

      2.times { sheet.rows[sales_b.length+3].add_cell nil }
      sheet.rows[sales_b.length+3].add_cell 'инстанти', style: fmt_tot
      sheet.rows[sales_b.length+3].add_cell "=SUM(J#{idx_inst + 3}:J#{sales_b.length + 2})",
        style: fmt_tot
      sheet.rows[sales_b.length+3].add_cell "=SUM(K#{idx_inst + 3}:K#{sales_b.length + 2})",
        style: fmt_tot
      sheet.rows[sales_b.length+3].add_cell nil

      2.times { sheet.rows[sales_b.length+4].add_cell nil }
      sheet.rows[sales_b.length+4].add_cell 'вкупно', style: fmt_tot
      sheet.rows[sales_b.length+4].add_cell "=SUM(J3:J#{sales_b.length + 2})",
        style: fmt_tot
      2.times { sheet.rows[sales_b.length+4].add_cell nil }
      # append empty rows if needed
      ([sales_b.length, sales_c.length].max - sales_b.length).times do |i|
        6.times { sheet.rows[sales_b.length+5+i].add_cell nil }
      end

      # C 
      idx_inst = sales_c.find_index {|s| s.is_instant == 1}
      sales_c.each_with_index do |s, idx|
        style = (idx+1 == sales_c.length) ? fmt_urow : fmt_row
        cols.each_with_index do |c, i|
          sheet.rows[idx + 2].add_cell s[c], style: style[i]
        end
      end
      # totals
      2.times { sheet.rows[sales_c.length+2].add_cell nil }
      sheet.rows[sales_c.length+2].add_cell 'лото', style: fmt_tot
      sheet.rows[sales_c.length+2].add_cell "=SUM(P3:P#{idx_inst + 2})",
        style: fmt_tot
      sheet.rows[sales_c.length+2].add_cell "=SUM(Q3:Q#{idx_inst + 2})",
        style: fmt_tot
      sheet.rows[sales_c.length+2].add_cell nil

      2.times { sheet.rows[sales_c.length+3].add_cell nil }
      sheet.rows[sales_c.length+3].add_cell 'инстанти', style: fmt_tot
      sheet.rows[sales_c.length+3].add_cell "=SUM(P#{idx_inst + 3}:P#{sales_c.length + 2})",
        style: fmt_tot
      sheet.rows[sales_c.length+3].add_cell "=SUM(Q#{idx_inst + 3}:Q#{sales_c.length + 2})",
        style: fmt_tot
      sheet.rows[sales_c.length+3].add_cell nil

      2.times { sheet.rows[sales_c.length+4].add_cell nil }
      sheet.rows[sales_c.length+4].add_cell 'вкупно', style: fmt_tot
      sheet.rows[sales_c.length+4].add_cell "=SUM(P3:P#{sales_c.length + 2})",
        style: fmt_tot
      2.times { sheet.rows[sales_c.length+4].add_cell nil }

      sheet.column_widths 3.8, 16, 8, 12, 10, 6,
        3.8, 16, 8, 12, 10, 6,
        3.8, 16, 8, 12, 10, 6
    end
  end # compare months

  ##
  # Create sheets for inactive terminals 
  # for instants games
  def self.x_create_inactive_sheets book, day
    instant_games = Sale.select('DISTINCT s.game_id AS game_id').
                      where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day).
                      where("g.type = 'INSTANT'").
                      where('s.sales <> 0.0'). # check if needed
                      joins('AS s INNER JOIN games AS g ON s.game_id = g.id').
                      order('s.game_id')
    instant_ids = instant_games.to_a.map { |r| r.game_id }

    terminals_with_instants = 
      Sale.select('DISTINCT terminal_id').
           where("date BETWEEN date(:day, 'start of month') AND :day", day: day).
           where('game_id IN (:instants)', instants: instant_ids).
           where('sales IS NOT NULL AND sales <> 0.0'). # here is needed
           order('terminal_id') 
    terminal_ids = terminals_with_instants.to_a.map { |r| r.terminal_id }

    # formating
    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_hdr_l = book.styles.add_style sz: 10,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_r = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 

    fmt_a = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '####', alignment: { horizontal: :center }
    fmt_b = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left }
    fmt_c = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }

    fmt_game = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      b: true, alignment: { horizontal: :center }

    # create worksheets
    summary = book.add_worksheet name: 'Инстанти-неак.тер.(сумарно)' do |sheet|
      sheet.add_row [
        "Терминали кои не примаат одреден инстант (сумарно)", '', '', ''],
        style: [fmt_merge]*3 + [nil], height: 32 
      sheet.merge_cells 'A1:C1'

      sheet.add_row [ 'id', 'инстант', "бр. неактивни\nтерминали"], 
        style: [fmt_hdr_c, fmt_hdr_l, fmt_hdr_r], height: 24
    end

    detailed = book.add_worksheet name: 'Инстанти-неак.тер.(детално)' do |sheet|
      sheet.add_row [
        "Терминали кои не примаат одреден инстант \n(детално)", '', ''],
        style: [fmt_merge]*2 + [nil], height: 32 
      sheet.merge_cells 'A1:B1'
    end

    # loop per instant
    det_row = 2
    instant_ids.each do |id|
      inactive_ids = 
        terminal_ids - 
        Terminal.select('t.id AS terminal_id, t.name AS name, SUM(s.sales) AS total').
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day).
          where('s.game_id = :game_id', game_id: id).
          joins('AS t INNER JOIN sales AS s ON t.id = s.terminal_id').
          group('t.id').
          having('total > 0.0').
          order('t.id').to_a.map { |r| r.terminal_id }

      # summary sheet
      summary.add_row [id, Game.find(id).name, inactive_ids.count],
        style: [fmt_a, fmt_b, fmt_c], height: 12
      
      # detailed sheet
      detailed.add_row [ Game.find(id).name ], style: [fmt_game]*1, height: 12
      detailed.merge_cells "A#{det_row}:B#{det_row}"
      det_row += 1
      inactive_ids.each do |id|
        detailed.add_row [id, Terminal.find(id).name], style: [fmt_a, fmt_b],
          height: 12
        det_row += 1
      end
    end
    summary.column_widths 3.8, 16, 12
    detailed.column_widths 10, 50
  end # inactive terminals
  
  def self.x_create_instants_terminals_sheet book, day
    instant_games = Sale
              .select('DISTINCT s.game_id AS game_id')
              .where("s.date BETWEEN date(:day, 'start of month') AND :day",
                day: day)
              .where("g.type = 'INSTANT'")
              .where('s.sales > 0.0') # check if needed
              .joins('AS s INNER JOIN games AS g ON s.game_id = g.id')
              .order('s.game_id')
    instant_ids = instant_games.to_a.map { |r| r.game_id }
    instant_names = { }
    instant_ids.each {|id| instant_names[id] = Game.find(id).name}

    qry =<<-EOT
      DISTINCT
        t.id        AS terminal_id,
        t.name      AS terminal_name,
        t.city      AS city,
        g.id        AS game_id,
        g.name      AS game_name
    EOT
    sales = Sale
            .select(qry)
            .joins('AS s INNER JOIN games AS g ON s.game_id = g.id')
            .joins('INNER JOIN terminals AS t ON s.terminal_id = t.id')
            .where("g.type = 'INSTANT'")
            .where("s.sales > 0")
            .where("date BETWEEN date(:day, 'start of month') AND :day",
              day: day)
            .order('t.id')
    arr = sales.to_a
    h   = { }
    arr.each do |r|
      # puts "#{r.terminal_id}, #{r.terminal_name}, #{r.game_name}"
      h[r.terminal_id] ||= {
        terminal_name: r.terminal_name,
        count: 0,
        city: r.city,
      }  
      h[r.terminal_id][r.game_id] = true
      h[r.terminal_id][:count] += 1
    end

    # formating
    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_hdr_l = book.styles.add_style sz: 10,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_r = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_inst = book.styles.add_style sz: 8,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 

    fmt_id = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '####', alignment: { horizontal: :center }
    fmt_term = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left }
    fmt_city = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :center }
    fmt_inst = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :center }

    fmt_uid = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '####', alignment: { horizontal: :center },
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_uterm = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_ucity = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :center },
      border: { style: :thin, edges: [:bottom], color: '000000' }
    fmt_uinst = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :center },
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_count = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', b: true,
      alignment: { horizontal: :center, vertical: :center }

    inst_arr = instant_ids.map {|id| instant_names[id]}
    book.add_worksheet name: 'Инстанти & Терминали' do |sheet|
      sheet.add_row [
        "Терминали кои не примаат одреден инстант во период" +
        " 01--#{day.strftime DMY}", '', '',] +
        ['']*(inst_arr.length),
        style: [fmt_merge]*(3+inst_arr.length), height: 32 
      sheet.merge_cells sheet.rows[0].cells[(0..inst_arr.length + 2)] 

      sheet.add_row [ 'id', 'терминал/прод. место', 'град'] + inst_arr, 
        style: [fmt_hdr_c, fmt_hdr_l, fmt_hdr_c] +
          [fmt_hdr_inst]*(inst_arr.length),
        height: 24
        
      h.delete_if {|k, v| v[:count] >= instant_ids.length }
      h.keys.sort.each_with_index do |terminal_id, idx|
        style = h.keys.length == idx+1 ?
          [fmt_uid, fmt_uterm, fmt_ucity] + [fmt_uinst]*(inst_arr.length) :
          [fmt_id, fmt_term, fmt_city] + [fmt_inst]*(inst_arr.length)
        arr = [ ]
        instant_ids.each {|id| arr.push(h[terminal_id][id] ? '✓' : '')}
        sheet.add_row [terminal_id, h[terminal_id][:terminal_name],
          h[terminal_id][:city]] + arr, height: 12,
          style: style
      end
      
      inst_col = ('A'.ord + 3).chr
      from = 3
      to   = h.keys.length + 2
      countblanks = [ ]
      inst_arr.length.times do |i|
        col = (inst_col.ord + i).chr
        countblanks.push "=COUNTBLANK(#{col}#{from}:#{col}#{to})"
      end
      sheet.add_row ['']*3 + countblanks, style: fmt_count

      sheet.column_widths 8, 32, 12, *[8]*(inst_arr.length)

      sheet.sheet_view.pane do |p|
        p.top_left_cell = 'A4'
        p.state = :frozen_split
        p.y_split = 2
        p.active_pane = :bottom_right
      end
    end
  end # instants/terminals
 
  ##
  # Create sheets for top terminal sales per game
  # 
  def self.x_create_top_terminals_sheet book, day, opt = {}
    # top 
    top_count   = opt[:top_count]
    top_count ||= TOP_COUNT

    total_sales = Sale
        .where("date BETWEEN date(:day, 'start of month') AND :day",
          day: day)
        .sum(:sales)

    game_sales = Sale
      .select('s.game_id AS game_id, g.name AS name, g.price AS price,' +
        ' SUM(s.sales) AS total_sales')
      .joins('AS s INNER JOIN games AS g ON s.game_id = g.id')
      .where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day)
      .where('s.sales > 0')
      .group('s.game_id')
      .order('s.game_id')

    # formating
    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true

    fmt_hdr_l = book.styles.add_style sz: 10,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_r = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_game_a = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '####', alignment: { horizontal: :center }, b: true, 
      border: { style: :dotted, edges: [:bottom], color: '000000' }
    fmt_game_b = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :left }, b: true, 
      border: { style: :dotted, edges: [:bottom], color: '000000' }
    fmt_game_c = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      alignment: { horizontal: :center }, format_code: '#,###', b: true, 
      border: { style: :dotted, edges: [:bottom], color: '000000' }
    fmt_game_d = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }, b: true, 
      border: { style: :dotted, edges: [:bottom], color: '000000' }
    fmt_game_e = book.styles.add_style sz: 10, font_name: 'Droid Sans',
      format_code: '#0.00%', alignment: { horizontal: :right }, b: true, 
      border: { style: :dotted, edges: [:bottom], color: '000000' }

    fmt_term_a = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '####', alignment: { horizontal: :center }
    fmt_term_b = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      alignment: { horizontal: :left }
    fmt_term_c = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      alignment: { horizontal: :center }, format_code: '#,###'
    fmt_term_d = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }
    fmt_term_e = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '#0.00%', alignment: { horizontal: :right }

    book.add_worksheet name: 'Топ терминали' do |sheet|
      # top rows
      sheet.add_row ["Период: #{ (day - (day.day - 1)).strftime DMY } --" +
        " #{ day.strftime DMY }"], style: fmt_merge, height: 18
      sheet.add_row ["Вкупна уплата: #{ thou_sep(total_sales.to_i) }"],
        style: fmt_merge, height: 18
      sheet.add_row [nil], style: fmt_merge, height: 18
      sheet.merge_cells 'A1:E1'
      sheet.merge_cells 'A2:E2'

      sheet.add_row [ "id", "игра/терминал", "цена/град", "промет",
        "учество %" ], style: [fmt_hdr_c, fmt_hdr_l, fmt_hdr_c, fmt_hdr_r,
        fmt_hdr_r], height: 24
  
      game_sales.each do |s|
        sheet.add_row [s.game_id, s.name, s.price, s.total_sales,
          s.total_sales*1.0/total_sales], style: [fmt_game_a, fmt_game_b,
          fmt_game_c, fmt_game_d, fmt_game_e], height: 14

        top_terminals = Sale
          .select('s.terminal_id AS terminal_id, t.name AS name,' +
                ' t.city AS city, SUM(s.sales) AS total_sales')
          .joins('AS s INNER JOIN terminals AS t ON s.terminal_id = t.id')
          .where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day)
          .where('s.game_id' => s.game_id)
          .where('s.sales > 0')
          .group('s.terminal_id')
          .order('total_sales DESC')
          .limit(top_count).each do |t|

            sheet.add_row [t.terminal_id, t.name, t.city, t.total_sales,
              t.total_sales*1.0/s.total_sales], style: [fmt_term_a, fmt_term_b,
              fmt_term_c, fmt_term_d, fmt_term_e], height: 12
          end
          
      end
      
      # selection
      sheet.sheet_view {|v| v.add_selection(:top_left, { active_cell: 'A3' })}

      # fix column widths
      sheet.column_widths 8, 42, 12, 10, 10

      sheet.sheet_view.pane do |p|
        p.top_left_cell = 'A4'
        p.state = :frozen_split
        p.y_split = 4
        p.active_pane = :bottom_right
      end

      # selection
      sheet.sheet_view {|v| v.add_selection(:bottom_right, 
        { active_cell: 'A3', sqref: 'A3' })}
    end
  end # top terminals

  ##
  # Create sheets for top terminal sales per game
  # 
  def self.x_create_sales_per_city_sheet book, day
    total_sales = Sale
      .where("date BETWEEN date(:day, 'start of month') AND :day", day: day)
      .sum(:sales)

    term_sales = Sale.
          select('s.terminal_id AS terminal_id, t.name AS name,' +
                 ' t.city AS city, SUM(s.sales) AS sales').
          joins('AS s INNER JOIN terminals AS t ON s.terminal_id = t.id').
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day).
          where('s.sales > 0.0').
          group('s.terminal_id').to_a

    qry =<<-EOT
      SELECT
        ts.city                         AS city,
        SUM(ts.sales)                   AS city_sales,
        COUNT(DISTINCT ts.terminal_id)  AS term_count,
        MIN(ts.sales)                   AS min_term_sales,
        AVG(ts.sales)                   AS avg_term_sales,
        MAX(ts.sales)                   AS max_term_sales
      FROM (
        SELECT
          s.terminal_id   AS terminal_id,
          t.city          AS city,
          SUM(s.sales)    AS sales
        FROM
          sales AS s 
          INNER JOIN terminals AS t
            ON s.terminal_id = t.id
        WHERE
          s.date BETWEEN date('#{ day.strftime YMD_FMT }', 'start of month') AND 
          '#{ day.strftime YMD_FMT }'
          AND s.sales > 0
        GROUP BY s.terminal_id
      ) AS ts
      GROUP BY ts.city
      ORDER BY term_count DESC
    EOT

    # ActiveRecord::Base.logger = Logger.new(STDOUT)
    city_sales = ActiveRecord::Base.connection.execute qry, day: day

    # formating
    fmt_merge = book.styles.add_style sz: 12,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true

    fmt_hdr_l = book.styles.add_style sz: 10,
      alignment: { horizontal: :left, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_c = book.styles.add_style sz: 10,
      alignment: { horizontal: :center, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' } 
    fmt_hdr_r = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom], color: '000000' }

    fmt_hdr_rd = book.styles.add_style sz: 10,
      alignment: { horizontal: :right, vertical: :center, wrap_text: true },
      font_name: 'Droid Sans', b: true, 
      border: { style: :thin, edges: [:bottom, :right], color: '000000' }
    book.styles do |s| # change right border from thin to dotted
      border = s.borders[s.cellXfs[fmt_hdr_rd].borderId]
      border.prs.each {|part| part.style = :dotted if part.name == :right }
    end

    fmt_city = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      alignment: { horizontal: :left }
    fmt_sales = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }
    fmt_share = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '#0.00%', alignment: { horizontal: :center }
    fmt_count = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :center }
    fmt_min = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right }
    fmt_avg = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right },
      border: { style: :dotted, edges: [:right], color: '000000' }
    fmt_max = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '#,###', alignment: { horizontal: :right },
      fg_color: '006600'
    fmt_id = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      format_code: '####', alignment: { horizontal: :center },
      fg_color: '006600'
    fmt_name = book.styles.add_style sz: 8, font_name: 'Droid Sans',
      alignment: { horizontal: :right }, fg_color: '006600'

    book.add_worksheet name: 'Продажба по градови' do |sheet|
      # top rows
      sheet.add_row ["Период: #{ (day - (day.day - 1)).strftime DMY } --" +
        " #{ day.strftime DMY }"], style: fmt_merge, height: 18
      sheet.add_row ["Вкупна уплата: #{ thou_sep(total_sales.to_i) }"],
        style: fmt_merge, height: 18
      sheet.add_row [nil], style: fmt_merge, height: 18
      sheet.merge_cells 'A1:I1'
      sheet.merge_cells 'A2:I2'
      
      sheet.add_row %W{град  продажба учество бр.\nтерм. мин. прос. макс.
        id терминал}, style: [fmt_hdr_l, fmt_hdr_r, fmt_hdr_c, fmt_hdr_c,
        fmt_hdr_r, fmt_hdr_rd, fmt_hdr_r, fmt_hdr_c, fmt_hdr_r], height: 24

      city_sales.each do |s|
        t = term_sales.select {|ts| ts['city'] == s['city'] and
          ts['sales'] >= s['max_term_sales']}[0]
        sheet.add_row [
          s['city'], s['city_sales'], s['city_sales']*1.0/total_sales,
          s['term_count'], s['min_term_sales'], s['avg_term_sales'],
          s['max_term_sales'], t['terminal_id'], t['name']],
          style: [fmt_city, fmt_sales, fmt_share, fmt_count, fmt_min,
            fmt_avg, fmt_max, fmt_id, fmt_name], height: 12
      end
      # fix column widths
      sheet.column_widths 20, 10, 8, 6, 8, 9, 9, 8, 35

      # freeze pane
      sheet.sheet_view.pane do |p|
        p.top_left_cell = 'A4'
        p.state = :frozen_split
        p.y_split = 4
        p.active_pane = :bottom_right
      end

      # selection
      sheet.sheet_view {|v| v.add_selection(:bottom_right, 
        { active_cell: 'A3', sqref: 'A3' })}
    end
  end # sales per city
end

__END__
TODO:
  - add new sheet: annual report: from jan 1st until day
  - check if can do pivot tables
    
