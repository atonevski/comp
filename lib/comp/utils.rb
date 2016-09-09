# encoding: UTF-8

require 'active_record'

class Sale < ActiveRecord::Base
  belongs_to :terminal
  belongs_to :game
end

class Terminal < ActiveRecord::Base
  has_many :sales
end

class Agent < ActiveRecord::Base
  has_many :terminals
end

class Game < ActiveRecord::Base

  self.inheritance_column = nil # the problem with the column named type

  has_many :sales
end

module Comp
  RM_PERC         = 0.04
  MPM_PERC        = 0.03
  MONTH_NAMES_MK  = [ '', 'јануари', 'февруари', 'март', 'април', 'мај', 'јуни',
                      'јули', 'август', 'септември', 'октомври', 'ноември', 'декември' ]
  GAME_TYPES_MK   = {
    'LOTTO'     => 'лото',
    'NUMBER'    => 'џокер',
    'BINGO'     => 'бинго',
    'INSTANT'   => 'инстант',
    'NEWSPAPER' => 'ср.вести',
    'TOTO'      => 'тото',
  }
  TOP_COUNT = 10

  def self.get_sales(c)
    sel =<<-EOT
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
    sales = Sale.select(sel).
          joins('AS s INNER JOIN games AS g ON s.game_id = g.id').
          joins('INNER JOIN terminals AS t ON s.terminal_id = t.id').
          where('s.date BETWEEN :from AND :to', from: c[:from], to: c[:to])
    exclude = if c[:criteria][:exclude] then 'NOT' else '' end
    if    c[:criteria][:key] =~ /agents?/
      sales = sales.where('t.agent_id ' + exclude + ' IN (:items)', 
                items: c[:criteria][:items])
    elsif c[:criteria][:key] =~ /terminals?/
      sales = sales.where('t.id ' + exclude + ' IN (:items)', 
                items: c[:criteria][:items])
    end
    sales = sales.group('g.id').having('qty > 0').order('is_instant, parent_id, g.id')
  end

  # return last day of month for given data
  def self.last_day_of_month(d)
    raise "#{ d } is not a date" unless d.instance_of? Date
    if d.month == 12
      (Date.new d.year + 1, 1, 1) - 1
   else
      (Date.new d.year, d.month + 1, 1) - 1
    end
  end

  #
  # Compare every 7th day
  #
  def self.month_ago(period)
    year = period[:from].year
    if period[:from].month == 1
      year  = year - 1
      month = 12
    else
      month = period[:from].month - 1
    end
    {
      from: Date.new(year, month, period[:from].day),
      to:   Date.new(year, month, period[:to].day)
    }
  end

  def self.year_ago(period)
    {
      from: Date.new(period[:from].year - 1, period[:from].month, period[:from].day),
      to:   Date.new(period[:to].year - 1, period[:to].month, period[:to].day)
    }
  end

  def self.create_compare_sheet(book, day)
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

    sheet = book.create_worksheet name: 'Споредба по 7 денa'

    # worksheet default format
    sheet.default_format = Spreadsheet::Format.new font: Spreadsheet::Font.new('Droid Sans')

    # Formating
    font      =  Spreadsheet::Font.new('Droid Sans') 
    font_sm   =  Spreadsheet::Font.new('Droid Sans', size: 8) 
    bold_font_sm =  Spreadsheet::Font.new('Droid Sans', size: 8, bold: true) 
    bold_font =  Spreadsheet::Font.new('Droid Sans', bold: true) 
    line_fmt  = [
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#0.00%',   font: font_sm),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#0.00%',   font: font_sm),
    ]
    lastln_fmt  = [
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :left, font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
        font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
        font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#0.00%',
        font: font_sm),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#0.00%',
        font: font_sm),
    ]
    heading_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              bottom: :thin, text_wrap: true, vertical_align: :middle, 
              font: Spreadsheet::Font.new('Droid Sans', bold: true)
    heading_fmt  = [
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :center, font: bold_font),
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :left, font: bold_font),
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :right, font: bold_font),
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :right, font: bold_font),
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :center, font: bold_font),
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :center, font: bold_font),
    ]
    total_fmt = Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',
                  font: bold_font)
    totperc_fmt = Spreadsheet::Format.new(horizontal_align: :right, 
                  number_format: '#0.00%', font: bold_font_sm)
    title_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              text_wrap: true, vertical_align: :middle, bottom: :thin, 
              font: Spreadsheet::Font.new('Droid Sans', bold: true, size: 12)

    sheet.column(1).width = sheet.column(8).width = sheet.column(13).width = 20
    sheet.column(0).width = sheet.column(7).width = sheet.column(12).width = 3
    sheet.column(2).width = sheet.column(9).width = sheet.column(14).width = 14
    sheet.column(3).width = sheet.column(10).width = sheet.column(15).width = 12
    
    sheet.row(0).height = 18
    sheet.row(1).height = 24
  
    # table A
    sheet.row(0)[0] = "А: #{ rounded_period[:from].strftime YMD_FMT } --" +
                      " #{ rounded_period[:to].strftime YMD_FMT }"
    0.upto(5) { |c| sheet.row(0).set_format c, title_fmt }
    sheet.merge_cells 0, 0, 0, 5

    sheet.row(1).push *[ "id", "игра", "пари", "тикети / комб.", "А vs Б", "А vs В" ]
    0.upto(5) { |i| sheet.row(1).set_format i, heading_fmt[i] }

    ri = 2
    sales_a.each_with_index do |s, idx|
      r = sheet.row ri
      r[0] = s.game_id
      r[1] = s.name
      r[2] = s.sales
      r[3] = s.qty
      
      b = sales_b.select {|r| r.game_id == s.game_id}[0]
      if b
        r[4] = (s.sales - b.sales)/b.sales
      end

      c = sales_c.select {|r| r.game_id == s.game_id}[0]
      if c
        r[5] = (s.sales - c.sales)/c.sales
      end

      ln_fmt = if idx == sales_a.length - 1 then lastln_fmt else line_fmt end
      0.upto(5) { |i| r.set_format i, ln_fmt[i] }
      
      ri += 1
    end
    # totals A
    compare_total_for sheet, sales_a, ri, 1
    row_a = ri
    
    # table B
    a_month_ago = month_ago rounded_period
    sheet.row(0)[7] = "Б: #{ a_month_ago[:from].strftime YMD_FMT } --" +
                      " #{ a_month_ago[:to].strftime YMD_FMT }"
    7.upto(10) { |c| sheet.row(0).set_format c, title_fmt }
    sheet.merge_cells 0, 7, 0, 10

    off = 7
    sheet.row(1)[off + 0] = "id"
    sheet.row(1)[off + 1] = "игра"
    sheet.row(1)[off + 2] = "пари"
    sheet.row(1)[off + 3] = "тикети / комб."
    0.upto(3) { |i| sheet.row(1).set_format off + i, heading_fmt[i] }

    ri = 2
    sales_b.each_with_index do |s, idx|
      r = sheet.row ri
      r[off + 0] = s.game_id
      r[off + 1] = s.name
      r[off + 2] = s.sales
      r[off + 3] = s.qty
      ln_fmt = if idx == sales_b.length - 1 then lastln_fmt else line_fmt end
      0.upto(3) { |i| r.set_format off + i, ln_fmt[i] }
      
      ri += 1
    end
    # totals B
    compare_total_for sheet, sales_b, ri, 8
    row_b = ri

    # table C
    a_year_ago = year_ago rounded_period
    sheet.row(0)[12] = "В: #{ a_year_ago[:from].strftime YMD_FMT } --" +
                      " #{ a_year_ago[:to].strftime YMD_FMT }"
    12.upto(15) { |c| sheet.row(0).set_format c, title_fmt }
    sheet.merge_cells 0, 12, 0, 15
    off = 12
    sheet.row(1)[off + 0] = "id"
    sheet.row(1)[off + 1] = "игра"
    sheet.row(1)[off + 2] = "пари"
    sheet.row(1)[off + 3] = "тикети / комб."
    0.upto(3) { |i| sheet.row(1).set_format off + i, heading_fmt[i] }

    ri = 2
    sales_c.each_with_index do |s, idx|
      r = sheet.row ri
      r[off + 0] = s.game_id
      r[off + 1] = s.name
      r[off + 2] = s.sales
      r[off + 3] = s.qty
      ln_fmt = if idx == sales_c.length - 1 then lastln_fmt else line_fmt end
      0.upto(3) { |i| r.set_format off + i, ln_fmt[i] }
      
      ri += 1
    end
    # totals C
    compare_total_for sheet, sales_c, ri, 13
    row_c = ri

    # increase AvB, AvC
    sheet.row(row_a)[4] = (sheet.row(row_a)[2] - sheet.row(row_b)[9]) / sheet.row(row_b)[9]
    sheet.row(row_a + 1)[4] = (sheet.row(row_a + 1)[2] - sheet.row(row_b + 1)[9]) / 
                              sheet.row(row_b + 1)[9]
    sheet.row(row_a + 2)[4] = (sheet.row(row_a + 2)[2] - sheet.row(row_b + 2)[9]) / 
                              sheet.row(row_b + 2)[9]
    sheet.row(row_a).set_format 4, totperc_fmt
    sheet.row(row_a + 1).set_format 4, totperc_fmt
    sheet.row(row_a + 2).set_format 4, totperc_fmt

    sheet.row(row_a)[5] = (sheet.row(row_a)[2] - sheet.row(row_c)[14]) / sheet.row(row_c)[14]
    sheet.row(row_a + 1)[5] = (sheet.row(row_a + 1)[2] - sheet.row(row_c + 1)[14]) / 
                              sheet.row(row_c + 1)[14]
    sheet.row(row_a + 2)[5] = (sheet.row(row_a + 2)[2] - sheet.row(row_c + 2)[14]) / 
                              sheet.row(row_c + 2)[14]
    sheet.row(row_a).set_format 5, totperc_fmt
    sheet.row(row_a + 1).set_format 5, totperc_fmt
    sheet.row(row_a + 2).set_format 5, totperc_fmt

  end

  def self.compare_total_for sheet, sales, row, col
    bold_font =  Spreadsheet::Font.new('droid sans', bold: true) 
    total_fmt = Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',
                  font: bold_font)

    lotto = { }
    lotto[:money] = sales.inject(0) { |sum, s|
          if s.is_instant != 1 then sum + s.sales else sum end }
    lotto[:columns] = sales.inject(0) { |sum, s|
          if s.is_instant != 1 then sum + s.qty else sum end }
    sheet.row(row)[col] = 'лото'
    sheet.row(row).set_format col, total_fmt
    sheet.row(row)[col + 1] = lotto[:money]
    sheet.row(row).set_format col + 1, total_fmt
    sheet.row(row)[col + 2] = lotto[:columns]
    sheet.row(row).set_format col + 2, total_fmt
    
    instants = { }
    instants[:money] = sales.inject(0) { |sum, s|
        if s.is_instant == 1 then sum + s.sales else sum end }
    instants[:tickets] = sales.inject(0) { |sum, s|
        if s.is_instant == 1 then sum + s.qty else sum end }
    sheet.row(row + 1)[col] = 'инстанти'
    sheet.row(row + 1).set_format col, total_fmt
    sheet.row(row + 1)[col + 1] = instants[:money]
    sheet.row(row + 1).set_format col + 1, total_fmt
    sheet.row(row + 1)[col + 2] = instants[:tickets]
    sheet.row(row + 1).set_format col + 2, total_fmt
    
    sheet.row(row + 2)[col] = 'вкупно'
    sheet.row(row + 2).set_format col, total_fmt
    sheet.row(row + 2)[col + 1] = lotto[:money] + instants[:money]
    sheet.row(row + 2).set_format col + 1, total_fmt
  end
  
  ##
  # WriteExel: compare sheet
  #
  def self.we_create_compare_sheet(book, day)
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

    sheet = book.add_worksheet 'Споредба по 7 денa'
    
    fmt_merge = book.add_format center_across: 1, bold: 1, size: 12, align: 'vcenter',
                  font: 'Droid Sans', bottom: 1, border_color: 'black'
    fmt_header = book.add_format font: 'Droid Sans', bold: 1, size: 10, align: 'vcenter',
                text_wrap: 1, bottom: 1, border_color: 'black'

    fmt_header_c = book.add_format font: 'Droid Sans', bold: 1, size: 10, align: 'vcenter',
                text_wrap: 1, bottom: 1, border_color: 'black'
    fmt_header_c.set_align('center')
    fmt_header_l = book.add_format font: 'Droid Sans', bold: 1, size: 10, align: 'vcenter',
                text_wrap: 1, bottom: 1, border_color: 'black'
    fmt_header_l.set_align('left')
    fmt_header_r = book.add_format font: 'Droid Sans', bold: 1, size: 10, align: 'vcenter',
                text_wrap: 1, bottom: 1, border_color: 'black'
    fmt_header_r.set_align('right')

    fmt_a = book.add_format font: 'Droid Sans', size: 10, align: 'center'
    fmt_b = book.add_format font: 'Droid Sans', size: 10, align: 'left'
    fmt_c = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'
    fmt_d = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'
    fmt_e = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#0.00%'
    fmt_f = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#0.00%'

    fmt_ua = book.add_format font: 'Droid Sans', size: 10, align: 'center',
             bottom: 1, border_color: 'black'
    fmt_ub = book.add_format font: 'Droid Sans', size: 10, align: 'left',
             bottom: 1, border_color: 'black'
    fmt_uc = book.add_format font: 'Droid Sans', size: 10, align: 'right', 
             num_format: '#,###', bottom: 1, border_color: 'black'
    fmt_ud = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#,###', bottom: 1, border_color: 'black'
    fmt_ue = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#0.00%', bottom: 1, border_color: 'black'
    fmt_uf = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#0.00%', bottom: 1, border_color: 'black'
    
    fmt_bb = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             bold: 1
    fmt_bc = book.add_format font: 'Droid Sans', size: 10, align: 'right', 
             num_format: '#,###', bold: 1
    fmt_bd = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#,###', bold: 1
    fmt_be = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#0.00%', bold: 1
    fmt_bf = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#0.00%', bold: 1

    # fix column widths
    sheet.set_row 0, 18
    sheet.set_row 1, 24
    sheet.set_column 'B:B', 20
    sheet.set_column 'I:I', 20
    sheet.set_column 'N:N', 20
    sheet.set_column 'A:A', 4
    sheet.set_column 'H:H', 4
    sheet.set_column 'M:M', 4
    sheet.set_column 'C:C', 14
    sheet.set_column 'J:J', 14
    sheet.set_column 'O:O', 14
    sheet.set_column 'D:D', 12
    sheet.set_column 'K:K', 12
    sheet.set_column 'P:P', 12

    # table A
    sheet.merge_range 'A1:F1', "А: #{ rounded_period[:from].strftime YMD_FMT } --" +
                      " #{ rounded_period[:to].strftime YMD_FMT }", fmt_merge
    sheet.write 'A2', 'id', fmt_header_c
    sheet.write 'B2', 'игра', fmt_header_l
    sheet.write 'C2', 'пари', fmt_header_r
    sheet.write 'D2', 'тикети / комб.', fmt_header_r
    sheet.write 'E2', 'А vs Б', fmt_header_c
    sheet.write 'F2', 'А vs В', fmt_header_c

    sales_a.each_with_index do |s, idx|
      sheet.write 2 + idx, 0, s.game_id,  (idx + 1 == sales_a.length ? fmt_ua : fmt_a)
      sheet.write 2 + idx, 1, s.name,     (idx + 1 == sales_a.length ? fmt_ub : fmt_b)
      sheet.write 2 + idx, 2, s.sales,    (idx + 1 == sales_a.length ? fmt_uc : fmt_c)
      sheet.write 2 + idx, 3, s.qty,      (idx + 1 == sales_a.length ? fmt_ud : fmt_d)

      # a vs b
      b = sales_b.find_index {|r| r.game_id == s.game_id}
      if b
        sheet.write 'E' + (idx + 3).to_s,
          "=(C" + (idx + 3).to_s + "-J" + (b + 3).to_s + ")/J" + (b + 3).to_s,
          (idx + 1 == sales_a.length ? fmt_ue : fmt_e)
      else
        sheet.write 'E' + (idx + 3).to_s, '', (idx + 1 == sales_a.length ? fmt_ue : fmt_e)
      end
      
      # a vs c
      c = sales_c.find_index {|r| r.game_id == s.game_id}
      if c
        sheet.write 'F' + (idx + 3).to_s,
          "=(C" + (idx + 3).to_s + "-O" + (c + 3).to_s + ")/O" + (c + 3).to_s,
          (idx + 1 == sales_a.length ? fmt_ue : fmt_e)
      else
        sheet.write 'F' + (idx + 3).to_s, '', (idx + 1 == sales_a.length ? fmt_uf : fmt_f)
      end
    end

    idx_instants = sales_a.find_index {|s| s.is_instant == 1} # first instant

    sheet.write 'B' + (sales_a.length + 3).to_s, 'лото', fmt_bb
    sheet.write 'C' + (sales_a.length + 3).to_s, 
          "=SUM(C3:C" + (idx_instants + 2).to_s + ")", fmt_bc
    sheet.write 'D' + (sales_a.length + 3).to_s, 
          "=SUM(D3:D" + (idx_instants + 2).to_s + ")", fmt_bd

    sheet.write 'B' + (sales_a.length + 4).to_s, 'инстанти', fmt_bb
    sheet.write 'C' + (sales_a.length + 4).to_s, 
          "=SUM(C" + (idx_instants + 3).to_s + ":C" + (sales_a.length + 2).to_s + ")", fmt_bc
    sheet.write 'D' + (sales_a.length + 4).to_s, 
          "=SUM(D" + (idx_instants + 3).to_s + ":D" + (sales_a.length + 2).to_s + ")", fmt_bd

    sheet.write 'B' + (sales_a.length + 5).to_s, 'вкупно', fmt_bb
    sheet.write 'C' + (sales_a.length + 5).to_s, 
          "=SUM(C3:C" + (sales_a.length + 2).to_s + ")", fmt_bc

    sheet.write 'E' + (sales_a.length + 3).to_s,
          "=(C" + (sales_a.length + 3).to_s + "-J" + (sales_b.length + 3).to_s + ")/J" +
            (sales_b.length + 3).to_s, fmt_be
    sheet.write 'E' + (sales_a.length + 4).to_s,
          "=(C" + (sales_a.length + 4).to_s + "-J" + (sales_b.length + 4).to_s + ")/J" +
            (sales_b.length + 4).to_s, fmt_be
    sheet.write 'E' + (sales_a.length + 5).to_s,
          "=(C" + (sales_a.length + 5).to_s + "-J" + (sales_b.length + 5).to_s + ")/J" +
            (sales_b.length + 5).to_s, fmt_be

    sheet.write 'F' + (sales_a.length + 3).to_s,
          "=(C" + (sales_a.length + 3).to_s + "-O" + (sales_c.length + 3).to_s + ")/O" +
            (sales_c.length + 3).to_s, fmt_be
    sheet.write 'F' + (sales_a.length + 4).to_s,
          "=(C" + (sales_a.length + 4).to_s + "-O" + (sales_c.length + 4).to_s + ")/O" +
            (sales_c.length + 4).to_s, fmt_be
    sheet.write 'F' + (sales_a.length + 5).to_s,
          "=(C" + (sales_a.length + 5).to_s + "-O" + (sales_c.length + 5).to_s + ")/O" +
            (sales_c.length + 5).to_s, fmt_be

    # table B
    a_month_ago = month_ago rounded_period
    sheet.merge_range 'H1:K1', "Б: #{ a_month_ago[:from].strftime YMD_FMT } --" +
                      " #{ a_month_ago[:to].strftime YMD_FMT }", fmt_merge
    sheet.write 'H2', 'id', fmt_header_c
    sheet.write 'I2', 'игра', fmt_header_l
    sheet.write 'J2', 'пари', fmt_header_r
    sheet.write 'K2', 'тикети / комб.', fmt_header_r

    sales_b.each_with_index do |s, idx|
      sheet.write 2 + idx,  7, s.game_id,  (idx + 1 == sales_b.length ? fmt_ua : fmt_a)
      sheet.write 2 + idx,  8, s.name,     (idx + 1 == sales_b.length ? fmt_ub : fmt_b)
      sheet.write 2 + idx,  9, s.sales,    (idx + 1 == sales_b.length ? fmt_uc : fmt_c)
      sheet.write 2 + idx, 10, s.qty,      (idx + 1 == sales_b.length ? fmt_ud : fmt_d)
    end
    idx_instants = sales_b.find_index {|s| s.is_instant == 1} # first instant

    sheet.write 'I' + (sales_b.length + 3).to_s, 'лото', fmt_bb
    sheet.write 'J' + (sales_b.length + 3).to_s, 
          "=SUM(J3:J" + (idx_instants + 2).to_s + ")", fmt_bc
    sheet.write 'K' + (sales_b.length + 3).to_s, 
          "=SUM(K3:K" + (idx_instants + 2).to_s + ")", fmt_bd

    sheet.write 'I' + (sales_b.length + 4).to_s, 'инстанти', fmt_bb
    sheet.write 'J' + (sales_b.length + 4).to_s, 
          "=SUM(J" + (idx_instants + 3).to_s + ":J" + (sales_b.length + 2).to_s + ")", fmt_bc
    sheet.write 'K' + (sales_b.length + 4).to_s, 
          "=SUM(K" + (idx_instants + 3).to_s + ":K" + (sales_b.length + 2).to_s + ")", fmt_bd

    sheet.write 'I' + (sales_b.length + 5).to_s, 'вкупно', fmt_bb
    sheet.write 'J' + (sales_b.length + 5).to_s, 
          "=SUM(J3:J" + (sales_b.length + 2).to_s + ")", fmt_bc

    # table C
    a_year_ago = year_ago rounded_period
    sheet.merge_range 'M1:P1', "В: #{ a_year_ago[:from].strftime YMD_FMT } --" +
                      " #{ a_year_ago[:to].strftime YMD_FMT }", fmt_merge
    sheet.write 'M2', 'id', fmt_header_c
    sheet.write 'N2', 'игра', fmt_header_l
    sheet.write 'O2', 'пари', fmt_header_r
    sheet.write 'P2', 'тикети / комб.', fmt_header_r

    sales_c.each_with_index do |s, idx|
      sheet.write 2 + idx, 12, s.game_id,  (idx + 1 == sales_c.length ? fmt_ua : fmt_a)
      sheet.write 2 + idx, 13, s.name,     (idx + 1 == sales_c.length ? fmt_ub : fmt_b)
      sheet.write 2 + idx, 14, s.sales,    (idx + 1 == sales_c.length ? fmt_uc : fmt_c)
      sheet.write 2 + idx, 15, s.qty,      (idx + 1 == sales_c.length ? fmt_ud : fmt_d)
    end
    idx_instants = sales_c.find_index {|s| s.is_instant == 1} # first instant

    sheet.write 'N' + (sales_c.length + 3).to_s, 'лото', fmt_bb
    sheet.write 'O' + (sales_c.length + 3).to_s, 
          "=SUM(O3:O" + (idx_instants + 2).to_s + ")", fmt_bc
    sheet.write 'P' + (sales_c.length + 3).to_s, 
          "=SUM(P3:P" + (idx_instants + 2).to_s + ")", fmt_bd

    sheet.write 'N' + (sales_c.length + 4).to_s, 'инстанти', fmt_bb
    sheet.write 'O' + (sales_c.length + 4).to_s, 
          "=SUM(O" + (idx_instants + 3).to_s + ":O" + (sales_c.length + 2).to_s + ")", fmt_bc
    sheet.write 'P' + (sales_c.length + 4).to_s, 
          "=SUM(P" + (idx_instants + 3).to_s + ":P" + (sales_c.length + 2).to_s + ")", fmt_bd

    sheet.write 'N' + (sales_c.length + 5).to_s, 'вкупно', fmt_bb
    sheet.write 'O' + (sales_c.length + 5).to_s, 
          "=SUM(O3:O" + (sales_c.length + 2).to_s + ")", fmt_bc

  end

  ##
  # Share sheet
  #
  def self.create_share_sheet book, day
    
    sheet = book.create_worksheet name: 'Учество'

    # worksheet default format
    sheet.default_format = Spreadsheet::Format.new font: Spreadsheet::Font.new('Droid Sans')

    # Formating
    font      =  Spreadsheet::Font.new('Droid Sans') 
    font_sm   =  Spreadsheet::Font.new('Droid Sans', size: 8) 
    bold_font_sm =  Spreadsheet::Font.new('Droid Sans', size: 8, bold: true) 
    bold_font =  Spreadsheet::Font.new('Droid Sans', bold: true) 
    line_fmt  = [
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#0.00%',   font: font_sm),
    ]
    lastln_fmt  = [
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :left, font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
        font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
        font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
        font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#0.00%',
        font: font_sm),
    ]
    totperc_fmt = Spreadsheet::Format.new(horizontal_align: :right, 
                  number_format: '#0.00%', font: bold_font_sm)
    bold_font = Spreadsheet::Font.new('droid sans', bold: true) 
    total_fmt = Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',
                  font: bold_font)
    title_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              text_wrap: true, vertical_align: :middle, bottom: :thin,
              font: Spreadsheet::Font.new('Droid Sans', bold: true, size: 12)

    sheet.column(0).width = sheet.column(7).width = 3
    sheet.column(1).width = sheet.column(8).width = 20
    sheet.column(2).width = sheet.column(9).width = 14
    sheet.column(3).width = sheet.column(10).width = 12

    sheet.row(0).height = 20
    
    sheet.row(1).push *[ "id", "игра", "пари", "тикети / комб.", "бр. терм.", "удел", '',
                         "id", "игра", "пари", "тикети / комб.", "бр. терм.", "удел"]
    sheet.row(1).height = 25
    heading_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              bottom: :thin, text_wrap: true, vertical_align: :middle, 
              font: Spreadsheet::Font.new('Droid Sans', bold: true)
    0.upto(5) { |i| sheet.row(1).set_format i, heading_fmt
                    sheet.row(1).set_format i + 7, heading_fmt }

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

    # indirect
    sheet.row(0)[0] = "Индиректна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }"
    0.upto(5) { |c| sheet.row(0).set_format c, title_fmt }
    sheet.merge_cells 0, 0, 0, 5
    sales_i.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[0] = s.game_id
      r[1] = s.name
      r[2] = s.sales
      r[3] = s.qty
      r[4] = s.term_count
      dir = sales_d.select { |ds| ds.game_id == s.game_id }[0]
      if dir
        r[5] = s.sales / (s.sales + dir.sales)
      end
      ln_fmt = if idx == sales_i.length - 1 then lastln_fmt else line_fmt end
      0.upto(5) { |i| r.set_format i, ln_fmt[i] }
    end
    compare_total_for sheet, sales_i, sales_i.length + 2, 1

    # direct
    sheet.row(0)[7] = "Директна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }"
    7.upto(12) { |c| sheet.row(0).set_format c, title_fmt }
    sheet.merge_cells 0, 7, 0, 12
    off = 7
    sales_d.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[off + 0] = s.game_id
      r[off + 1] = s.name
      r[off + 2] = s.sales
      r[off + 3] = s.qty
      r[off + 4] = s.term_count
      ind = sales_i.select { |is| is.game_id == s.game_id }[0]
      if ind
        r[off + 5] = s.sales / (s.sales + ind.sales)
      end
      ln_fmt = if idx == sales_d.length - 1 then lastln_fmt else line_fmt end
      0.upto(5) { |i| r.set_format off + i, ln_fmt[i] }
    end
    compare_total_for sheet, sales_d, sales_d.length + 2, 8

    # total share
    row_i = sales_i.length + 2
    row_d = sales_d.length + 2
    sheet.row(row_i)[5] = sheet.row(row_i)[2] / (sheet.row(row_i)[2] + sheet.row(row_d)[9])
    sheet.row(row_i + 1)[5] = sheet.row(row_i + 1)[2] / 
          (sheet.row(row_i + 1)[2] + sheet.row(row_d + 1)[9])
    sheet.row(row_i + 2)[5] = sheet.row(row_i + 2)[2] / 
          (sheet.row(row_i + 2)[2] + sheet.row(row_d + 2)[9])
    sheet.row(row_i).set_format 5, totperc_fmt
    sheet.row(row_i + 1).set_format 5, totperc_fmt
    sheet.row(row_i + 2).set_format 5, totperc_fmt

    sheet.row(row_d)[12] = sheet.row(row_d)[9] / (sheet.row(row_i)[2] + sheet.row(row_d)[9])
    sheet.row(row_d + 1)[12] = sheet.row(row_d + 1)[9] / 
          (sheet.row(row_i + 1)[2] + sheet.row(row_d + 1)[9])
    sheet.row(row_d + 2)[12] = sheet.row(row_d + 2)[9] / 
          (sheet.row(row_i + 2)[2] + sheet.row(row_d + 2)[9])
    sheet.row(row_d).set_format 12, totperc_fmt
    sheet.row(row_d + 1).set_format 12, totperc_fmt
    sheet.row(row_d + 2).set_format 12, totperc_fmt

    qry =<<-EOT
      CASE 
        WHEN g.type = 'INSTANT' THEN 1
        ELSE 0
      END                     AS is_instant,
      SUM(s.sales)            AS sales,
      COUNT(DISTINCT t.id)    AS term_count
    EOT
    sales = Sale.select(qry).
          joins('AS s INNER JOIN games AS g ON s.game_id = g.id').
          joins('INNER JOIN terminals AS t ON s.terminal_id = t.id')
    term_i = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day). 
          where('t.agent_id <> 225').
          where('s.sales > 0').
          group('is_instant').
          order('is_instant')
    term_d = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day). 
          where('t.agent_id = 225').
          where('s.sales > 0').
          group('is_instant').
          order('is_instant')

    qry =<<-EOT
      SUM(s.sales)            AS sales,
      COUNT(DISTINCT t.id)    AS term_count
    EOT
    sales = Sale.select(qry).
          joins('AS s INNER JOIN terminals AS t ON s.terminal_id = t.id')
    totterm_i = sales.
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day). 
          where('t.agent_id <> 225').
          where('s.sales > 0')
    totterm_d = sales.
          where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day). 
          where('t.agent_id = 225').
          where('s.sales > 0')

    sheet.row(row_i + 0)[4] = term_i[0].term_count
    sheet.row(row_i + 1)[4] = term_i[1].term_count
    sheet.row(row_i + 2)[4] = totterm_i[0].term_count
    sheet.row(row_i + 0).set_format 4, total_fmt
    sheet.row(row_i + 1).set_format 4, total_fmt
    sheet.row(row_i + 2).set_format 4, total_fmt

    sheet.row(row_d + 0)[11] = term_d[0].term_count
    sheet.row(row_d + 1)[11] = term_d[1].term_count
    sheet.row(row_d + 2)[11] = totterm_d[0].term_count
    sheet.row(row_d + 0).set_format 11, total_fmt
    sheet.row(row_d + 1).set_format 11, total_fmt
    sheet.row(row_d + 2).set_format 11, total_fmt

    # google chart
    total = sheet.row(row_i+0)[2].to_i + sheet.row(row_i+1)[2].to_i +
            sheet.row(row_d+0)[9].to_i + sheet.row(row_d+1)[9].to_i

    p = sprintf "%5.2f%%|%5.2f%%|%5.2f%%|%5.2f%%", sheet.row(row_i+0)[2]/total*100,
                sheet.row(row_i+1)[2]/total*100,
                sheet.row(row_d+0)[9]/total*100,
                sheet.row(row_d+1)[9]/total*100

    sheet.column(14).default_format = line_fmt[1]
    sheet.column(15).default_format = line_fmt[2]

    r = sheet.row(2)
    r[14] = 'индиректна лото'
    r[15] = sheet.row(row_i + 0)[2]

    r = sheet.row(3)
    r[14] = 'индиректна инстанти'
    r[15] = sheet.row(row_i + 1)[2]

    r = sheet.row(4)
    r[14] = 'директна лото'
    r[15] = sheet.row(row_d + 0)[9]

    r = sheet.row(5)
    r[14] = 'директна инстанти'
    r[15] = sheet.row(row_d + 1)[9]
  end

  ##
  # WriteExel: compare sheet
  #
  def self.we_create_share_sheet book, day
    sheet = book.add_worksheet 'Sheet2', true
    
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

    fmt_merge = book.add_format center_across: 1, bold: 1, size: 12, align: 'vcenter',
                  font: 'Droid Sans', bottom: 1, border_color: 'black'

    # fix column widths and row heights
    sheet.set_row 0, 18
    sheet.set_row 1, 24
    sheet.set_column 'A:A', 3
    sheet.set_column 'H:H', 3
    sheet.set_column 'B:B', 20
    sheet.set_column 'I:I', 20
    sheet.set_column 'C:C', 14
    sheet.set_column 'J:J', 14
    sheet.set_column 'D:D', 12
    sheet.set_column 'K:K', 12

    fmt_merge = book.add_format center_across: 1, bold: 1, size: 12, align: 'vcenter',
                  font: 'Droid Sans', bottom: 1, border_color: 'black'
    fmt_header = book.add_format font: 'Droid Sans', bold: 1, size: 10, align: 'vcenter',
                text_wrap: 1, bottom: 1, border_color: 'black'

    fmt_header_c = book.add_format font: 'Droid Sans', bold: 1, size: 10, align: 'vcenter',
                text_wrap: 1, bottom: 1, border_color: 'black'
    fmt_header_c.set_align('center')
    fmt_header_l = book.add_format font: 'Droid Sans', bold: 1, size: 10, align: 'vcenter',
                text_wrap: 1, bottom: 1, border_color: 'black'
    fmt_header_l.set_align('left')
    fmt_header_r = book.add_format font: 'Droid Sans', bold: 1, size: 10, align: 'vcenter',
                text_wrap: 1, bottom: 1, border_color: 'black'
    fmt_header_r.set_align('right')

    fmt_a = book.add_format font: 'Droid Sans', size: 10, align: 'center'
    fmt_b = book.add_format font: 'Droid Sans', size: 10, align: 'left'
    fmt_c = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'
    fmt_d = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'
    fmt_e = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'
    fmt_f = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#0.00%'

    fmt_ua = book.add_format font: 'Droid Sans', size: 10, align: 'center',
             bottom: 1, border_color: 'black'
    fmt_ub = book.add_format font: 'Droid Sans', size: 10, align: 'left',
             bottom: 1, border_color: 'black'
    fmt_uc = book.add_format font: 'Droid Sans', size: 10, align: 'right', 
             num_format: '#,###', bottom: 1, border_color: 'black'
    fmt_ud = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#,###', bottom: 1, border_color: 'black'
    fmt_ue = book.add_format font: 'Droid Sans', size: 10, align: 'center',
             num_format: '#,###', bottom: 1, border_color: 'black'
    fmt_uf = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#0.00%', bottom: 1, border_color: 'black'
    
    fmt_bb = book.add_format font: 'Droid Sans', size: 10, align: 'right', bold: 1
    fmt_bc = book.add_format font: 'Droid Sans', size: 10, align: 'right', 
             num_format: '#,###', bold: 1
    fmt_bd = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#,###', bold: 1
    fmt_be = book.add_format font: 'Droid Sans', size: 10, align: 'center',
             num_format: '#,###', bold: 1
    fmt_bf = book.add_format font: 'Droid Sans', size: 10, align: 'right',
             num_format: '#0.00%', bold: 1

    fmt_o = book.add_format font: 'Droid Sans', size: 10, align: 'left'
    fmt_p = book.add_format font: 'Droid Sans', size: 10, align: 'right', hidden: 1,
            num_format: '#,###'

    # indirect
    sheet.merge_range 'A1:F1', "Индиректна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }",
      fmt_merge
    sheet.write 'A2', 'id', fmt_header_c
    sheet.write 'B2', 'игра', fmt_header_l
    sheet.write 'C2', 'пари', fmt_header_r
    sheet.write 'D2', 'тикети / комб.', fmt_header_r
    sheet.write 'E2', "бр. терм.", fmt_header_c
    sheet.write 'F2', "удел", fmt_header_c

    sales_i.each_with_index do |s, idx|
      sheet.write 2 + idx, 0, s.game_id,    (idx + 1 == sales_i.length ? fmt_ua : fmt_a)
      sheet.write 2 + idx, 1, s.name,       (idx + 1 == sales_i.length ? fmt_ub : fmt_b)
      sheet.write 2 + idx, 2, s.sales,      (idx + 1 == sales_i.length ? fmt_uc : fmt_c)
      sheet.write 2 + idx, 3, s.qty,        (idx + 1 == sales_i.length ? fmt_ud : fmt_d)
      sheet.write 2 + idx, 4, s.term_count, (idx + 1 == sales_i.length ? fmt_ud : fmt_e)

      # ind vs dir 
      d = sales_d.find_index {|r| r.game_id == s.game_id}
      if d
        sheet.write 'F' + (idx + 3).to_s,
          "=C" + (idx + 3).to_s + "/(J" + (d + 3).to_s + "+C" + (idx + 3).to_s + ")",
          (idx + 1 == sales_i.length ? fmt_uf : fmt_f)
      else
        sheet.write 'F' + (idx + 3).to_s, '', (idx + 1 == sales_i.length ? fmt_ue : fmt_e)
      end
    end
    idx_instants = sales_i.find_index {|s| s.is_instant == 1} # first instant

    sheet.write 'B' + (sales_i.length + 3).to_s, 'лото', fmt_bb
    sheet.write 'C' + (sales_i.length + 3).to_s, 
          "=SUM(C3:C" + (idx_instants + 2).to_s + ")", fmt_bc
    sheet.write 'D' + (sales_i.length + 3).to_s, 
          "=SUM(D3:D" + (idx_instants + 2).to_s + ")", fmt_bd
    sheet.write 'F' + (sales_i.length + 3).to_s, 
          "=C" + (sales_i.length + 3).to_s + "/(C" + (sales_i.length + 3).to_s + " + J" +
          (sales_d.length + 3).to_s + ")", fmt_bf

    sheet.write 'B' + (sales_i.length + 4).to_s, 'инстанти', fmt_bb
    sheet.write 'C' + (sales_i.length + 4).to_s, 
          "=SUM(C" + (idx_instants + 3).to_s + ":C" + (sales_i.length + 2).to_s + ")", fmt_bc
    sheet.write 'D' + (sales_i.length + 4).to_s, 
          "=SUM(D" + (idx_instants + 3).to_s + ":D" + (sales_i.length + 2).to_s + ")", fmt_bd
    sheet.write 'F' + (sales_i.length + 4).to_s, 
          "=C" + (sales_i.length + 4).to_s + "/(C" + (sales_i.length + 4).to_s + " + J" +
          (sales_d.length + 4).to_s + ")", fmt_bf

    sheet.write 'B' + (sales_i.length + 5).to_s, 'вкупно', fmt_bb
    sheet.write 'C' + (sales_i.length + 5).to_s, 
          "=SUM(C3:C" + (sales_i.length + 2).to_s + ")", fmt_bc
    sheet.write 'F' + (sales_i.length + 5).to_s, 
          "=C" + (sales_i.length + 5).to_s + "/(C" + (sales_i.length + 5).to_s + " + J" +
          (sales_d.length + 5).to_s + ")", fmt_bf

    # direct
    sheet.merge_range 'H1:M1', "Директна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }",
      fmt_merge
    sheet.write 'H2', 'id', fmt_header_c
    sheet.write 'I2', 'игра', fmt_header_l
    sheet.write 'J2', 'пари', fmt_header_r
    sheet.write 'K2', 'тикети / комб.', fmt_header_r
    sheet.write 'L2', "бр. терм.", fmt_header_c
    sheet.write 'M2', "удел", fmt_header_c
    sales_d.each_with_index do |s, idx|
      sheet.write 2 + idx,  7, s.game_id,    (idx + 1 == sales_d.length ? fmt_ua : fmt_a)
      sheet.write 2 + idx,  8, s.name,       (idx + 1 == sales_d.length ? fmt_ub : fmt_b)
      sheet.write 2 + idx,  9, s.sales,      (idx + 1 == sales_d.length ? fmt_uc : fmt_c)
      sheet.write 2 + idx, 10, s.qty,        (idx + 1 == sales_d.length ? fmt_ud : fmt_d)
      sheet.write 2 + idx, 11, s.term_count, (idx + 1 == sales_d.length ? fmt_ud : fmt_e)
      #
      # ind vs dir 
      i = sales_i.find_index {|r| r.game_id == s.game_id}
      if i
        sheet.write 'M' + (idx + 3).to_s,
          "=J" + (idx + 3).to_s + "/(J" + (i + 3).to_s + "+C" + (idx + 3).to_s + ")",
          (idx + 1 == sales_i.length ? fmt_uf : fmt_f)
      else
        sheet.write 'M' + (idx + 3).to_s, '', (idx + 1 == sales_i.length ? fmt_ue : fmt_e)
      end
    end
    idx_instants = sales_d.find_index {|s| s.is_instant == 1} # first instant

    sheet.write 'I' + (sales_d.length + 3).to_s, 'лото', fmt_bb
    sheet.write 'J' + (sales_d.length + 3).to_s, 
          "=SUM(J3:J" + (idx_instants + 2).to_s + ")", fmt_bc
    sheet.write 'K' + (sales_d.length + 3).to_s, 
          "=SUM(K3:K" + (idx_instants + 2).to_s + ")", fmt_bd
    sheet.write 'M' + (sales_i.length + 3).to_s, 
          "=J" + (sales_d.length + 3).to_s + "/(C" + (sales_i.length + 3).to_s + " + J" +
          (sales_d.length + 3).to_s + ")", fmt_bf

    sheet.write 'I' + (sales_d.length + 4).to_s, 'инстанти', fmt_bb
    sheet.write 'J' + (sales_d.length + 4).to_s, 
          "=SUM(J" + (idx_instants + 3).to_s + ":J" + (sales_d.length + 2).to_s + ")", fmt_bc
    sheet.write 'K' + (sales_d.length + 4).to_s, 
          "=SUM(K" + (idx_instants + 3).to_s + ":K" + (sales_d.length + 2).to_s + ")", fmt_bd
    sheet.write 'M' + (sales_i.length + 4).to_s, 
          "=J" + (sales_d.length + 4).to_s + "/(C" + (sales_i.length + 4).to_s + " + J" +
          (sales_d.length + 4).to_s + ")", fmt_bf

    sheet.write 'I' + (sales_d.length + 5).to_s, 'вкупно', fmt_bb
    sheet.write 'J' + (sales_d.length + 5).to_s, 
          "=SUM(J3:J" + (sales_d.length + 2).to_s + ")", fmt_bc
    sheet.write 'M' + (sales_i.length + 5).to_s, 
          "=J" + (sales_d.length + 5).to_s + "/(C" + (sales_i.length + 5).to_s + " + J" +
          (sales_d.length + 5).to_s + ")", fmt_bf

    sheet.write 'O2', 'Учество во вкупен промет', fmt_o
    sheet.write 'O3', 'индиректна лото', fmt_o
    sheet.write 'O4', 'индиректна инстанти', fmt_o
    sheet.write 'O5', 'директна лото', fmt_o
    sheet.write 'O6', 'директна инстанти', fmt_o
    sheet.write 'P3', '=Sheet2!C' + (sales_i.length + 3).to_s, fmt_p
    sheet.write 'P4', '=C' + (sales_i.length + 4).to_s, fmt_p
    sheet.write 'P5', '=J' + (sales_d.length + 3).to_s, fmt_p
    sheet.write 'P6', '=J' + (sales_d.length + 4).to_s, fmt_p

    puts sheet.name
    puts sheet.name.encoding

    chart = book.add_chart type: 'Chart::Pie', embedded: true , name_utf16be: true
    chart.set_title name: 'Учество во вкупен промет'
    chart.add_series categories: "=Sheet2!$O$3:$O$6",
       values: "=Sheet2!$P$3:$P$6"
    chart.set_legend position: 'right'

    sheet.insert_chart 'A' + (sales_i.length+7).to_s, chart
  end

  ##
  # Remainder sheet
  #
  def self.create_remainder_sheet book, day

    funds_file      = File.expand_path '../../../config/funds.yml', __FILE__
    commission_file = File.expand_path '../../../config/commission.yml', __FILE__

    funds       = YAML::load File.open funds_file
    commission  = YAML::load File.open commission_file

    # Formating
    font      =  Spreadsheet::Font.new('Droid Sans') 
    font_sm   =  Spreadsheet::Font.new('Droid Sans', size: 8) 
    bold_font_sm =  Spreadsheet::Font.new('Droid Sans', size: 8, bold: true) 
    bold_font =  Spreadsheet::Font.new('Droid Sans', bold: true) 
    line_i_fmt  = [
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(horizontal_align: :center, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :center, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#0.00%',  font: font_sm),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#0.00%',  font: font_sm),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
    ]
    lastln_i_fmt  = [
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#0.00%',
          font: font_sm),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#0.00%',
          font: font_sm),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
          font: font),
    ]
    line_d_fmt  = [
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(horizontal_align: :center, number_format: '#,###',  font: font),
      Spreadsheet::Format.new(horizontal_align: :center, number_format: '#,###',  font: font_sm),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#0.00%',  font: font_sm),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
    ]
    lastln_d_fmt  = [
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, number_format: '#,###',
          font: font_sm),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#0.00%',
          font: font_sm),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
          font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
          font: font),
    ]
    total_fmt = Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',
                  font: bold_font)
    totperc_fmt = Spreadsheet::Format.new(horizontal_align: :right, 
                  number_format: '#0.00%', font: bold_font_sm)
    title_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              text_wrap: true, vertical_align: :middle, bottom: :thin,
              font: Spreadsheet::Font.new('Droid Sans', bold: true, size: 12)

    sheet = book.create_worksheet name: 'Остаток'

    # worksheet default format
    sheet.default_format = Spreadsheet::Format.new font: Spreadsheet::Font.new('Droid Sans')

    sheet.column(0).width = sheet.column(11).width = 3
    sheet.column(1).width = sheet.column(12).width = 20
    sheet.column(2).width = sheet.column(13).width = 6
    sheet.column(3).width = sheet.column(14).width = 6
    sheet.column(5).width = 6
    sheet.column(6).width = 12
    sheet.column(8).width = 12
    sheet.column(9).width = sheet.column(19).width = 11


    sheet.row(0).height = 20
    sheet.row(1).height = 34
    sheet.row(1).push *[ "id", "игра", "цена", "терм.", "фонд за доб.", "пров.", "уплата",
                        "тикети / комб.", "фонд + пров. + МПМ + РМ", "остаток", "",
                        "id", "игра", "цена", "терм.", "фонд за доб.", "уплата",
                        "тикети / комб.", "фонд + РМ", "остаток",
                       ]
    heading_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              bottom: :thin, text_wrap: true, vertical_align: :middle, 
              font: Spreadsheet::Font.new('Droid Sans', bold: true)
    0.upto(9) { |i| sheet.row(1).set_format i, heading_fmt }
    0.upto(8) { |i| sheet.row(1).set_format 11 + i, heading_fmt }

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

    # indirect
    sheet.row(0)[0] = "Индиректна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }"
    0.upto(9) { |c| sheet.row(0).set_format c, title_fmt }
    sheet.merge_cells 0, 0, 0, 9
    r8_tot = 0
    r9_tot = 0
    sales_i.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[0] = s.game_id
      r[1] = s.name
      r[2] = s.price
      r[3] = s.term_count
      raise "Undefined funds for #{ s.game_id }" unless funds[s.game_id.to_s]
      r[4] = funds[s.game_id.to_s]
      raise "Undefined commission for #{ s.game_id }" unless commission[s.game_id.to_s]
      r[5] = commission[s.game_id.to_s]
      r[6] = s.sales
      r[7] = s.qty
      r[8] = (r[4] + r[5] + RM_PERC + MPM_PERC) * s.sales
      r[9] = s.sales - r[8]

      # update totals
      r8_tot += r[8]
      r9_tot += r[9]

      ln_fmt = if idx == sales_i.length - 1 then lastln_i_fmt else line_i_fmt end
      0.upto(9) { |i| r.set_format i, ln_fmt[i] }
    end
    row = sales_i.length
    sheet.row(row + 2)[6] = sales_i.inject(0) { |sum, s| sum + s.sales }
    sheet.row(row + 2).set_format 6, total_fmt
    sheet.row(row + 2)[8] = r8_tot
    sheet.row(row + 2).set_format 8, total_fmt
    sheet.row(row + 2)[9] = r9_tot
    sheet.row(row + 2).set_format 9, total_fmt
    sheet.row(row + 3)[9] = r9_tot / sheet.row(row + 2)[6]
    sheet.row(row + 3).set_format 9, totperc_fmt

    # direct
    sheet.row(0)[11] = "Директна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }"
    11.upto(19) { |c| sheet.row(0).set_format 11, title_fmt }
    sheet.merge_cells 0, 11, 0, 19
    r7_tot  = 0
    r8_tot  = 0
    off     = 11
    sales_d.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[off + 0] = s.game_id
      r[off + 1] = s.name
      r[off + 2] = s.price
      r[off + 3] = s.term_count
      raise "Undefined funds for #{ s.game_id }" unless funds[s.game_id.to_s]
      r[off + 4] = funds[s.game_id.to_s]
      r[off + 5] = s.sales
      r[off + 6] = s.qty
      r[off + 7] = (r[off + 4] + RM_PERC) * s.sales
      r[off + 8] = s.sales - r[off + 7]

      # update totals
      r7_tot += r[off + 7]
      r8_tot += r[off + 8]

      ln_fmt = if idx == sales_d.length - 1 then lastln_d_fmt else line_d_fmt end
      0.upto(8) { |i| r.set_format off + i, ln_fmt[i] }
    end
    row = sales_d.length
    sheet.row(row + 2)[off + 5] = sales_d.inject(0) { |sum, s| sum + s.sales }
    sheet.row(row + 2).set_format off + 5, total_fmt
    sheet.row(row + 2)[off + 7] = r7_tot
    sheet.row(row + 2).set_format off + 7, total_fmt
    sheet.row(row + 2)[off + 8] = r8_tot
    sheet.row(row + 2).set_format off + 8, total_fmt
    sheet.row(row + 3)[off + 8] = r8_tot / sheet.row(row + 2)[off + 5]
    sheet.row(row + 3).set_format off + 8, totperc_fmt

    #
    # google chart indirect
    #

    # chart colors
    cat10_colors = %W{ 1f77b4 ff7f0e 2ca02c d62728 9467bd 8c564b e377c2 7f7f7f bcbd22 17becf } 

    # inderect
    arr_i = [ # funds, commission, MPM, RM
      sales_i.inject(0) { |sum, s| sum + s.sales * funds[s.game_id.to_s] },
      sales_i.inject(0) { |sum, s| sum + s.sales * commission[s.game_id.to_s] },
      sales_i.inject(0) { |sum, s| sum + s.sales * MPM_PERC },
      sales_i.inject(0) { |sum, s| sum + s.sales * RM_PERC },
    ]
    arr_i << (sheet.row(sales_i.length + 2)[6] - arr_i.inject(0) { |sum, s| sum + s })
    tot_i = arr_i.inject(0) { |sum, s| sum + s }
    p = sprintf "%5.2f%%|%5.2f%%|%5.2f%%|%5.2f%%|%5.2f%%", arr_i[0]/tot_i*100,
                arr_i[1]/tot_i*100,
                arr_i[2]/tot_i*100,
                arr_i[3]/tot_i*100,
                arr_i[4]/tot_i*100 # remainder
    s = "https://chart.googleapis.com/chart" +
        "?chs=720x400" + # chart size
        "&chma=10,50,40,10" + # chart margins
        "&cht=p" + # 2d pie chart
        "&chco=#{ cat10_colors[0..3].join ',' }" + # series color
        "&chd=t:#{ arr_i.map {|x| x.to_i}.join ',' }" +
        "&chds=a" +
        "&chxt=x" + # axis labels
        "&chxs=0,000000,14" + # axis label styles y:axis, N number *=prefix end, p=percent
        "&chtt=Структура на индиректна продажба" + # chart title
        "&chts=000000,18,c" +
        "&chl=#{ p }" +
        "&chdl=фонд за добивки|провизија на заст.|МПМ|РМ|остаток"
    sheet.row(sales_i.length + 4)[1] = Spreadsheet::Link.new URI.escape(s),
      "Структура на индиректна продажба", ''

    sheet.column(21).default_format = line_i_fmt[1]
    sheet.column(22).default_format = line_i_fmt[6]

    r = sheet.row(2)
    r[21] = 'фонд'
    r[22] = arr_i[0]

    r = sheet.row(3)
    r[21] = 'провизија'
    r[22] = arr_i[1]

    r = sheet.row(4)
    r[21] = 'МПМ'
    r[22] = arr_i[2]

    r = sheet.row(5)
    r[21] = 'РМ'
    r[22] = arr_i[3]

    r = sheet.row(6)
    r[21] = 'остаток'
    r[22] = arr_i[4]

    # derect
    arr_d = [ # funds, commission, MPM, RM
      sales_d.inject(0) { |sum, s| sum + s.sales * funds[s.game_id.to_s] },
      sales_d.inject(0) { |sum, s| sum + s.sales * RM_PERC },
    ]

    arr_d << (sheet.row(sales_d.length + 2)[off + 5] - arr_d.inject(0) { |sum, s| sum + s })
    tot_d = arr_d.inject(0) { |sum, s| sum + s }
    p = sprintf "%5.2f%%|%5.2f%%|%5.2f%%", arr_d[0]/tot_d*100,
                arr_d[1]/tot_d*100,
                arr_d[2]/tot_d*100
    s = "https://chart.googleapis.com/chart" +
        "?chs=720x400" + # chart size
        "&chma=10,60,40,10" + # chart margins
        "&cht=p" + # 2d pie chart
        "&chco=#{ cat10_colors[0..2].join ',' }" + # series color
        "&chd=t:#{ arr_d.map {|x| x.to_i}.join ',' }" +
        "&chds=a" +
        "&chxt=x" + # axis labels
        "&chxs=0,000000,14" + # axis label styles y:axis, N number *=prefix end, p=percent
        "&chtt=Структура на директна продажба" + # chart title
        "&chts=000000,18,c" +
        "&chl=#{ p }" +
        "&chdl=фонд за добивки|РМ|остаток"
    sheet.row(sales_d.length + 4)[off + 1] = Spreadsheet::Link.new URI.escape(s),
      "Структура на директна продажба", ''

    sheet.column(24).default_format = line_i_fmt[1]
    sheet.column(25).default_format = line_i_fmt[6]

    r = sheet.row(2)
    r[24] = 'фонд'
    r[25] = arr_d[0]

    r = sheet.row(3)
    r[24] = 'РМ'
    r[25] = arr_d[1]

    r = sheet.row(4)
    r[24] = 'остаток'
    r[25] = arr_d[2]

  end # remainder

  ##
  # Writeexcel: Remainder sheet
  #
  def self.we_create_remainder_sheet book, day
    funds_file      = File.expand_path '../../../config/funds.yml', __FILE__
    commission_file = File.expand_path '../../../config/commission.yml', __FILE__

    funds       = YAML::load File.open funds_file
    commission  = YAML::load File.open commission_file

    sheet = book.add_worksheet 'rr', true

    # put queries and table values here
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

    # formats
    fmt_merge = book.add_format center_across: 1, bold: 1, size: 12, valign: 'vcenter',
                  font: 'Droid Sans', bottom: 1, border_color: 'black'

    fmt_hdr_c = book.add_format font: 'Droid Sans', bold: 1, size: 10, valign: 'vcenter',
                align: 'center', text_wrap: 1, bottom: 1, border_color: 'black'
    fmt_hdr_l = book.add_format font: 'Droid Sans', bold: 1, size: 10, valign: 'vcenter',
                align: 'left', text_wrap: 1, bottom: 1, border_color: 'black'
    fmt_hdr_r = book.add_format font: 'Droid Sans', bold: 1, size: 10, valign: 'vcenter',
                align: 'right', text_wrap: 1, bottom: 1, border_color: 'black'

    fmt_a = book.add_format font: 'Droid Sans', size: 10, align: 'center'
    fmt_b = book.add_format font: 'Droid Sans', size: 10, align: 'left'
    fmt_c = book.add_format font: 'Droid Sans', size: 10, align: 'center', num_format: '#,###'
    fmt_d = book.add_format font: 'Droid Sans', size: 10, align: 'center', num_format: '#,###'
    fmt_e = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#0.00%'
    fmt_f = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#0.00%'
    fmt_g = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'
    fmt_h = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'
    fmt_i = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'
    fmt_j = book.add_format font: 'Droid Sans', size: 10, align: 'right', num_format: '#,###'

    fmt_ua = book.add_format font: 'Droid Sans', size: 10, align: 'center', bottom: 1
    fmt_ub = book.add_format font: 'Droid Sans', size: 10, align: 'left', bottom: 1
    fmt_uc = book.add_format font: 'Droid Sans', size: 10, align: 'center',
              num_format: '#,###', bottom: 1
    fmt_ud = book.add_format font: 'Droid Sans', size: 10, align: 'center',
              num_format: '#,###',  bottom: 1
    fmt_ue = book.add_format font: 'Droid Sans', size: 10, align: 'right',
              num_format: '#0.00%', bottom: 1
    fmt_uf = book.add_format font: 'Droid Sans', size: 10, align: 'right',
              num_format: '#0.00%', bottom: 1
    fmt_ug = book.add_format font: 'Droid Sans', size: 10, align: 'right',
              num_format: '#,###', bottom: 1
    fmt_uh = book.add_format font: 'Droid Sans', size: 10, align: 'right',
              num_format: '#,###', bottom: 1
    fmt_ui = book.add_format font: 'Droid Sans', size: 10, align: 'right',
              num_format: '#,###', bottom: 1
    fmt_uj = book.add_format font: 'Droid Sans', size: 10, align: 'right',
              num_format: '#,###', bottom: 1

    fmt_tot = book.add_format font: 'Droid Sans', size: 10, align: 'right',
              num_format: '#,###', bold: 1
    fmt_tp  = book.add_format font: 'Droid Sans', size: 10, align: 'right',
              num_format: '#0.00%', bold: 1

    fmt_v = book.add_format font: 'Droid Sans', size: 10, align: 'left'
    fmt_w = book.add_format font: 'Droid Sans', size: 10, align: 'right', hidden: 1,
            num_format: '#,###'
    # fix column widths and row heights
    sheet.set_row 0, 20
    sheet.set_row 1, 34

    sheet.set_column 'A:A', 4
    sheet.set_column 'B:B', 20
    sheet.set_column 'C:C', 6
    sheet.set_column 'D:D', 6
    sheet.set_column 'E:E', 10
    sheet.set_column 'F:F', 6
    sheet.set_column 'G:G', 12
    sheet.set_column 'H:H', 12
    sheet.set_column 'I:I', 12
    sheet.set_column 'J:J', 12

    sheet.set_column 'L:L', 4
    sheet.set_column 'M:M', 20
    sheet.set_column 'N:N', 6
    sheet.set_column 'O:O', 6
    sheet.set_column 'P:P', 10
    sheet.set_column 'Q:Q', 12
    sheet.set_column 'R:R', 12
    sheet.set_column 'S:S', 12
    sheet.set_column 'T:T', 12

    fmt_merge = book.add_format center_across: 1, bold: 1, size: 12, align: 'vcenter',
                  font: 'Droid Sans', bottom: 1, border_color: 'black'

    # ind
    sheet.merge_range 'A1:J1', "Индиректна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }",
      fmt_merge
    sheet.write 'A2', 'id',             fmt_hdr_c
    sheet.write 'B2', 'игра',           fmt_hdr_l
    sheet.write 'C2', 'цена',           fmt_hdr_c
    sheet.write 'D2', 'терм.',          fmt_hdr_c
    sheet.write 'E2', "фонд за доб.",   fmt_hdr_c
    sheet.write 'F2', "пров.",          fmt_hdr_c
    sheet.write 'G2', "уплата",         fmt_hdr_r
    sheet.write 'H2', "тикети / комб.", fmt_hdr_r
    sheet.write 'I2', "фонд + пров. + МПМ + РМ", fmt_hdr_c
    sheet.write 'J2', "остаток",        fmt_hdr_c

    sales_i.each_with_index do |s, idx|
      sheet.write 2 + idx, 0, s.game_id,    (idx + 1 == sales_i.length ? fmt_ua : fmt_a)
      sheet.write 2 + idx, 1, s.name,       (idx + 1 == sales_i.length ? fmt_ub : fmt_b)
      sheet.write 2 + idx, 2, s.price,      (idx + 1 == sales_i.length ? fmt_uc : fmt_c)
      sheet.write 2 + idx, 3, s.term_count, (idx + 1 == sales_i.length ? fmt_ud : fmt_d)
      raise "Undefined funds for #{ s.game_id }" unless funds[s.game_id.to_s]
      sheet.write 2 + idx, 4, funds[s.game_id.to_s],
                                            (idx + 1 == sales_i.length ? fmt_ue : fmt_e)
      raise "Undefined commission for #{ s.game_id }" unless commission[s.game_id.to_s]
      sheet.write 2 + idx, 5, commission[s.game_id.to_s],
                                            (idx + 1 == sales_i.length ? fmt_uf : fmt_f)

      sheet.write 2 + idx, 6, s.sales,      (idx + 1 == sales_i.length ? fmt_ug : fmt_g)
      sheet.write 2 + idx, 7, s.qty,        (idx + 1 == sales_i.length ? fmt_uh : fmt_h)
      sheet.write 2 + idx, 8, '=(E' + (idx + 3).to_s + '+ F' + (idx + 3).to_s + 
                              "+ #{ RM_PERC } + #{ MPM_PERC })*G" + (idx + 3).to_s,
                                            (idx + 1 == sales_i.length ? fmt_ui : fmt_i)
      sheet.write 2 + idx, 9, '=G' + (idx + 3).to_s + '- I' + (idx + 3).to_s,
                                            (idx + 1 == sales_i.length ? fmt_uj : fmt_j)
    end
    sheet.write 'G' + (sales_i.length + 3).to_s,
        '=SUM(G3:G' + (sales_i.length + 2).to_s + ')', fmt_tot
    sheet.write 'I' + (sales_i.length + 3).to_s,
        '=SUM(I3:I' + (sales_i.length + 2).to_s + ')', fmt_tot
    sheet.write 'J' + (sales_i.length + 3).to_s,
        '=SUM(J3:J' + (sales_i.length + 2).to_s + ')', fmt_tot
    sheet.write 'J' + (sales_i.length + 4).to_s,
        '=J' + (sales_i.length + 3).to_s + '/G' + (sales_i.length + 3).to_s, fmt_tp


    # dir
    sheet.merge_range 'L1:T1', "Директна продажба, #{ MONTH_NAMES_MK[day.month] } #{ day.year }",
      fmt_merge
    sheet.write 'L2', 'id',             fmt_hdr_c
    sheet.write 'M2', 'игра',           fmt_hdr_l
    sheet.write 'N2', 'цена',           fmt_hdr_c
    sheet.write 'O2', 'терм.',          fmt_hdr_c
    sheet.write 'P2', "фонд за доб.",   fmt_hdr_c
    sheet.write 'Q2', "уплата",         fmt_hdr_r
    sheet.write 'R2', "тикети / комб.", fmt_hdr_r
    sheet.write 'S2', "фонд + РМ",      fmt_hdr_c
    sheet.write 'T2', "остаток",        fmt_hdr_c

    sales_d.each_with_index do |s, idx|
      sheet.write 2 + idx, 11, s.game_id,   (idx + 1 == sales_d.length ? fmt_ua : fmt_a)
      sheet.write 2 + idx, 12, s.name,      (idx + 1 == sales_d.length ? fmt_ub : fmt_b)
      sheet.write 2 + idx, 13, s.price,     (idx + 1 == sales_d.length ? fmt_uc : fmt_c)
      sheet.write 2 + idx, 14, s.term_count,(idx + 1 == sales_d.length ? fmt_ud : fmt_d)
      raise "Undefined funds for #{ s.game_id }" unless funds[s.game_id.to_s]
      sheet.write 2 + idx, 15, funds[s.game_id.to_s],
                                            (idx + 1 == sales_d.length ? fmt_ue : fmt_e)
      sheet.write 2 + idx, 16, s.sales,     (idx + 1 == sales_d.length ? fmt_ug : fmt_g)
      sheet.write 2 + idx, 17, s.qty,       (idx + 1 == sales_d.length ? fmt_uh : fmt_h)
      sheet.write 2 + idx, 18, '=(P' + (idx + 3).to_s + "+ #{ RM_PERC })*Q" + (idx + 3).to_s,
                                            (idx + 1 == sales_d.length ? fmt_ui : fmt_i)
      sheet.write 2 + idx, 19, '=Q' + (idx + 3).to_s + '- S' + (idx + 3).to_s,
                                            (idx + 1 == sales_d.length ? fmt_uj : fmt_j)
    end
    sheet.write 'Q' + (sales_d.length + 3).to_s,
        '=SUM(Q3:Q' + (sales_d.length + 2).to_s + ')', fmt_tot
    sheet.write 'S' + (sales_d.length + 3).to_s,
        '=SUM(S3:S' + (sales_d.length + 2).to_s + ')', fmt_tot
    sheet.write 'T' + (sales_d.length + 3).to_s,
        '=SUM(T3:T' + (sales_d.length + 2).to_s + ')', fmt_tot
    sheet.write 'T' + (sales_d.length + 4).to_s,
        '=T' + (sales_i.length + 3).to_s + '/Q' + (sales_i.length + 3).to_s, fmt_tp

    # char i
    sheet.write 'V2', 'Структура на индиректна продажба', fmt_v
    sheet.write 'V3', 'фонд',                             fmt_v
    sheet.write 'V4', 'провизија',                        fmt_v
    sheet.write 'V5', 'МПМ',                              fmt_v
    sheet.write 'V6', 'РМ',                               fmt_v
    sheet.write 'V7', 'остаток',                          fmt_v

    sheet.write 'W3', sales_i.inject(0){|s, r| s + r.sales*funds[r.game_id.to_s]}, fmt_w
    sheet.write 'W4', sales_i.inject(0){|s, r| s + r.sales*commission[r.game_id.to_s]}, fmt_w
    sheet.write 'W5', '=G' + (sales_i.length + 3).to_s + " * #{ MPM_PERC }", fmt_w
    sheet.write 'W6', '=G' + (sales_i.length + 3).to_s + " * #{ RM_PERC }", fmt_w
    sheet.write 'W7', '=J' + (sales_i.length + 3).to_s, fmt_w

    chart_i = book.add_chart type: 'Chart::Pie', embedded: true, name_utf16be: true
    chart_i.set_title  name: 'Структура на индиректна продажба'
    chart_i.add_series categories: "=rr!$V$3:$V$7", values:  "=rr!$W$3:$W$7"
    chart_i.set_legend position: 'right'
    sheet.insert_chart 'A' + (sales_i.length+7).to_s, chart_i

    # char d
    sheet.write 'Y2', 'Структура на директна продажба', fmt_v
    sheet.write 'Y3', 'фонд',                           fmt_v
    sheet.write 'Y4', 'РМ',                             fmt_v
    sheet.write 'Y5', 'остаток',                        fmt_v

    sheet.write 'Z3', sales_d.inject(0){|s, r| s + r.sales*funds[r.game_id.to_s]}, fmt_w
    sheet.write 'Z4', '=Q' + (sales_i.length + 3).to_s + " * #{ RM_PERC }", fmt_w
    sheet.write 'Z5', '=T' + (sales_i.length + 3).to_s, fmt_w

    chart_d = book.add_chart type: 'Chart::Pie', embedded: true, name_utf16be: true
    chart_d.set_title  name: 'Структура на директна продажба'
    chart_d.add_series categories: "=rr!$Y$3:$Y$5", values:  "=rr!$Z$3:$Z$5"
    chart_d.set_legend position: 'right'
    sheet.insert_chart 'L' + (sales_d.length+7).to_s, chart_d

  end

  ##
  # Weekly sales
  #
  def self.create_weekly_sheet(book, day)
    sheet = book.create_worksheet name: 'Неделно'

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

    # worksheet default format
    sheet.default_format = Spreadsheet::Format.new font: Spreadsheet::Font.new('Droid Sans')
    
    # format date columns
    date_fmt_mk = Spreadsheet::Format.new number_format: 'DD.MM.YYYY', 
                  font: Spreadsheet::Font.new('Droid Sans')
    sheet.column(1).default_format = date_fmt_mk
    sheet.column(2).default_format = date_fmt_mk
    sheet.column(6).default_format = date_fmt_mk
    sheet.column(7).default_format = date_fmt_mk

    # format week number columns
    week_number_fmt = Spreadsheet::Format.new horizontal_align: :center, number_format: '00',
                      font: Spreadsheet::Font.new('Droid Sans')
    sheet.column(0).default_format = week_number_fmt
    sheet.column(5).default_format = week_number_fmt

    # sales format
    sales_fmt = Spreadsheet::Format.new horizontal_align: :right, number_format: '#,###',
                font: Spreadsheet::Font.new('Droid Sans')
    sheet.column(3).default_format = sales_fmt
    sheet.column(8).default_format = sales_fmt

    # heading format
    heading_fmt = Spreadsheet::Format.new horizontal_align: :center, bottom: :thin, 
                  font: Spreadsheet::Font.new('Droid Sans', bold: true)
    sheet.row(1).push *[ "седмица", "понед.", "недела", "уплата", "",
                         "седмица", "понед.", "недела", "уплата", "",
                       ]
    [0, 1, 2, 3, 5, 6, 7, 8].each do |c|
      sheet.row(1).set_format c, heading_fmt
    end

    # title format
    title_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              text_wrap: true, vertical_align: :middle, bottom: :thin,
              font: Spreadsheet::Font.new('Droid Sans', bold: true, size: 12)
    sheet.row(0).height = 20
    sheet.row(0)[0] = "Неделна продажба #{ day.year }"
    0.upto(3) { |c| sheet.row(0).set_format c, title_fmt }
    sheet.merge_cells 0, 0, 0, 3
    sheet.row(0)[5] = "Неделна продажба #{ day.year - 1 }"
    0.upto(8) { |c| sheet.row(0).set_format c, title_fmt }
    sheet.merge_cells 0, 5, 0, 8

    wsales_now.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[0] = s.week_number
      r[1] = Date.parse s.monday
      r[2] = Date.parse s.sunday
      r[3] = s.sales
    end

    off = 5
    wsales_before.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[off + 0] = s.week_number
      r[off + 1] = Date.parse s.monday
      r[off + 2] = Date.parse s.sunday
      r[off + 3] = s.sales
    end
    
    # series: comma separated
    ser_now     = wsales_now.map { |s| s.sales.to_i }.join(',')
    ser_before  = wsales_before.map { |s| s.sales.to_i }.join(',')

    # google chart
    s = "https://chart.googleapis.com/chart" +
        "?chs=720x400" + # chart size
        "&chma=60,30,30,30" + # chart margins
        "&cht=lc" + # line chart
        "&chco=e0440e,f6c7b6" + # series color
        "&chd=t:#{ ser_now }|#{ ser_before }" +
        "&chxt=x,y" + # axis labels
        "&chxs=1,N*s" + # axis label styles y:axis, N number *=prefix end, s=separated
        "&chxl=0:|#{ (0 .. last_week).to_a.join('|') }" +
        "&chls=3|2" + # line style (thickness)
        "&chds=a" +
        "&chtt=Споредба на неделен промет" + # chart title
        "&chts=000000,18,c" + # chart style: color, size, align
        "&chg=-1,-1,2,4" + # grid by axis tics
        "&chxtc=0,5" + # tic style
        "&chdlp=t" + # legend position
        "&chdl=#{ day.year }|#{ day.year - 1 }" # chart legend
    sheet.row(last_week + 5)[1] = Spreadsheet::Link.new URI.escape(s),
      "График: Споредба на неделен промет", ''
  end # weekly sales

  ##
  # Compare months sheet
  def Comp.create_compare_months_sheet book, day
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
          where("s.date BETWEEN date(:day, 'start of month') AND date(:day,'start of month','+1 month','-1 day')", day: day). 
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')

    # B:
    sales_b = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND date(:day,'start of month','+1 month','-1 day')", day: day_b). 
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')

    # C:
    sales_c = sales.      
          where("s.date BETWEEN date(:day, 'start of month') AND date(:day,'start of month','+1 month','-1 day')", day: day_c). 
          group('g.id').
          having('qty > 0').
          order('is_instant, parent_id, g.id')

    sheet = book.create_worksheet name: 'Споредба на месеци' 

    # worksheet default format
    sheet.default_format = Spreadsheet::Format.new font: Spreadsheet::Font.new('Droid Sans')

    # default foramts for 
    sales_fmt = Spreadsheet::Format.new horizontal_align: :right, number_format: '#,###',
                  font: Spreadsheet::Font.new('Droid Sans')
    sheet.column( 3).default_format = sales_fmt
    sheet.column( 4).default_format = sales_fmt
    sheet.column( 9).default_format = sales_fmt
    sheet.column(10).default_format = sales_fmt
    sheet.column(15).default_format = sales_fmt
    sheet.column(16).default_format = sales_fmt

    # fix some columns widths
    sheet.column(0).width = sheet.column(6).width = sheet.column(12).width = 3
    sheet.column(1).width = sheet.column(7).width = sheet.column(13).width = 20

    # title format
    title_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              text_wrap: true, vertical_align: :middle, bottom: :thin,
              font: Spreadsheet::Font.new('Droid Sans', bold: true, size: 12)
    sheet.row(0).height = 20
    sheet.row(0)[0] = "#{ MONTH_NAMES_MK[day.month] } #{ day.year }"
    sheet.row(0).set_format 0, title_fmt
    sheet.merge_cells 0, 0, 0, 4
    sheet.row(0)[6] = "#{ MONTH_NAMES_MK[day_b.month] } #{ day_b.year }"
    sheet.row(0).set_format 6, title_fmt
    sheet.merge_cells 0, 6, 0, 10
    sheet.row(0)[12] = "#{ MONTH_NAMES_MK[day_c.month] } #{ day_c.year }"
    sheet.row(0).set_format 12, title_fmt
    sheet.merge_cells 0, 12, 0, 16

    heading_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
              bottom: :thin, text_wrap: true, vertical_align: :middle, 
              font: Spreadsheet::Font.new('Droid Sans', bold: true)
    sheet.row(1).push *[ "id", "игра", "бр.\nтерм.", "уплата", "комб./\nтикети.", "",
                         "id", "игра", "бр.\nтерм.", "уплата", "комб./\nтикети.", "",
                         "id", "игра", "бр.\nтерм.", "уплата", "комб./\nтикети.", "",
                       ]
    0.upto(4) do |c|
      sheet.row(1).set_format c     , heading_fmt
      sheet.row(1).set_format c +  6, heading_fmt
      sheet.row(1).set_format c + 12, heading_fmt
    end
    sheet.row(1).height = 24

    # Formating
    font      =  Spreadsheet::Font.new('Droid Sans') 
    bold_font =  Spreadsheet::Font.new('Droid Sans', bold: true) 
    line_fmt  = [
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(horizontal_align: :center, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format:  '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format:  '#,###',   font: font),
    ]
    lastln_fmt  = [
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :center, number_format: '#,###',
        font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
        font: font),
      Spreadsheet::Format.new(bottom: :thin, horizontal_align: :right, number_format: '#,###',
        font: font),
    ]

    total_fmt = Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',
                  font: bold_font)

    sales_a.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[0] = s.game_id
      r[1] = s.name
      r[2] = s.term_count
      r[3] = s.sales
      r[4] = s.qty

      ln_fmt = if idx == sales_a.length - 1 then lastln_fmt else line_fmt end
      0.upto(4) { |i| r.set_format i, ln_fmt[i] }
    end
    sheet.row(2 + sales_a.length)[3] = sales_a.inject(0) {|sum, s| sum + s.sales}
    sheet.row(2 + sales_a.length)[4] = sales_a.inject(0) {|sum, s| sum + s.qty}
    sheet.row(2 + sales_a.length).set_format 3, total_fmt
    sheet.row(2 + sales_a.length).set_format 4, total_fmt

    off = 6
    sales_b.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[off + 0] = s.game_id
      r[off + 1] = s.name
      r[off + 2] = s.term_count
      r[off + 3] = s.sales
      r[off + 4] = s.qty

      ln_fmt = if idx == sales_b.length - 1 then lastln_fmt else line_fmt end
      0.upto(4) { |i| r.set_format off + i, ln_fmt[i] }
    end
    sheet.row(2 + sales_b.length)[off + 3] = sales_b.inject(0) {|sum, s| sum + s.sales}
    sheet.row(2 + sales_b.length)[off + 4] = sales_b.inject(0) {|sum, s| sum + s.qty}
    sheet.row(2 + sales_b.length).set_format off + 3, total_fmt
    sheet.row(2 + sales_b.length).set_format off + 4, total_fmt

    off = 12
    sales_c.each_with_index do |s, idx|
      r = sheet.row 2 + idx
      r[off + 0] = s.game_id
      r[off + 1] = s.name
      r[off + 2] = s.term_count
      r[off + 3] = s.sales
      r[off + 4] = s.qty

      ln_fmt = if idx == sales_c.length - 1 then lastln_fmt else line_fmt end
      0.upto(4) { |i| r.set_format off + i, ln_fmt[i] }
    end
    sheet.row(2 + sales_c.length)[off + 3] = sales_c.inject(0) {|sum, s| sum + s.sales}
    sheet.row(2 + sales_c.length)[off + 4] = sales_c.inject(0) {|sum, s| sum + s.qty}
    sheet.row(2 + sales_c.length).set_format off + 3, total_fmt
    sheet.row(2 + sales_c.length).set_format off + 4, total_fmt
  end

  ##
  # monthly sales 
  #
  def self.create_monthly_sheet book, day
    # IMPORTANT:
    # shuld limit dates up to full month!

    if day != last_day_of_month(day)
      day = (Date.new day.year, day.month, 1) - 1
    end

    qry =<<-EOT
      date(s.date, 'start of month')  AS fdom,
      SUM(s.sales)                    AS sales,
      COUNT(DISTINCT t.id)            AS term_count
    EOT
    sales = Sale.select(qry).
          joins('AS s INNER JOIN terminals AS t ON s.terminal_id = t.id').
          where('s.date <= :day', day: day).
          group('fdom').
          order('fdom')

    sheet = book.create_worksheet name: 'Месечно'

    # worksheet default format
    sheet.default_format = Spreadsheet::Format.new font: Spreadsheet::Font.new('Droid Sans')

    # fix some column widths
    sheet.column(0).width = 12
    sheet.column(2).width = 14

    # heading format
    heading_fmt = Spreadsheet::Format.new horizontal_align: :center, bottom: :thin, 
                  font: Spreadsheet::Font.new('Droid Sans', bold: true)
    sheet.row(1).push *[ "месец 'год.", "бр. терм.", "уплата", ]
    [0, 1, 2].each do |c|
      sheet.row(1).set_format c, heading_fmt
    end

    sales.each_with_index do |s, idx|
      r = sheet.row(2 + idx)
      r[0] = Date.parse s.fdom
      r[1] = s.term_count
      r[2] = s.sales

      r.set_format 0, Spreadsheet::Format.new(number_format: "MMM 'YY", align: :left,
                font: Spreadsheet::Font.new('Droid Sans'))
      r.set_format 1, Spreadsheet::Format.new(horizontal_align: :center, number_format: '#,###',
                font: Spreadsheet::Font.new('Droid Sans'))
      r.set_format 2, Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',
                font: Spreadsheet::Font.new('Droid Sans'))
    end

    # table for the pivot table: monthly sales per game type
    qry =<<-EOT
      SELECT
        s.fdom        AS month,
        s.game_type   AS type,
        SUM(s.sales)  AS sales,
        SUM(s.qty)    AS qty
      FROM (
        SELECT
          date(s.date, 'start of month')
                        AS fdom,
          g.id          AS game_id,
          g.name        AS game_name,
          CASE 
            WHEN g.parent IS NOT NULL THEN g.parent
            ELSE g.id
          END           AS parent_id,
          g2.type       AS game_type,
          s.sales       AS sales,
          s.sales/g.price
                        AS  qty
        FROM
          games AS g
          INNER JOIN games AS g2
            ON g2.id = parent_id
          INNER JOIN sales AS s
            ON g.id = s.game_id
        WHERE sales <> 0.0
      ) AS s
      GROUP BY s.fdom, s.game_type
      ORDER BY s.fdom, s.game_type
    EOT

    sales_by_type = ActiveRecord::Base.connection.execute qry
    sheet = book.create_worksheet name: 'monthly-by-type'
    
    # worksheet default format
    sheet.default_format = Spreadsheet::Format.new font: Spreadsheet::Font.new('Droid Sans')

    sheet.column(0).default_format = Spreadsheet::Format.new(number_format: "MMM 'YY", align: :left,
                    font: Spreadsheet::Font.new('Droid Sans'))
    sheet.column(1).default_format = Spreadsheet::Format.new(align: :left, 
                    font: Spreadsheet::Font.new('Droid Sans'))
    sheet.column(2).default_format = Spreadsheet::Format.new(number_format: "#,###", align: :right,
                    font: Spreadsheet::Font.new('Droid Sans'))
    sheet.column(3).default_format = Spreadsheet::Format.new(number_format: "#,###", align: :right,
                    font: Spreadsheet::Font.new('Droid Sans'))
    sheet.row(0).push *['month', 'type', 'sales', 'qty'] 
    sales_by_type.each_with_index do |s, idx|
      r    = sheet.row idx + 1
      r[0] = Date.parse s['month']
      r[1] = GAME_TYPES_MK[s['type']]
      r[2] = s['sales'].to_i
      r[3] = s['qty'].to_i
    end

  end # create_monthly_sheet

  ##
  # Create sheets for inactive terminals 
  # for instants games
  def self.create_inactive_sheets book, day
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

    # create the worksheet
    summary   = book.create_worksheet name: 'Инстанти - Неактивни терминали (сумарно)'
    detailed  = book.create_worksheet name: 'Инстанти - Неактивни терминали (детално)'
    
    # row indexes in each sheet
    sidx = 0
    didx = 0
    
    # insert titles in each sheet, heading and columns and rows widths and heights
    title_fmt = Spreadsheet::Format.new weight: :bold, horizontal_align: :center, 
                  text_wrap: true, vertical_align: :middle, bottom: :thin, 
                  font: Spreadsheet::Font.new('Droid Sans', bold: true, size: 12)
    font      =  Spreadsheet::Font.new('Droid Sans') 
    bold_font =  Spreadsheet::Font.new('Droid Sans', bold: true) 
    heading_fmt  = [
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :center, font: bold_font),
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :left, font: bold_font),
      Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :right, font: bold_font),
    ]
    sline_fmt  = [
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :left, font: font),
      Spreadsheet::Format.new(horizontal_align: :right, font: font),
    ]
    summary.column(0).width = 3
    summary.column(1).width = 25
    summary.column(2).width = 16
    summary.row(sidx)[0] = "Терминали кои не примаат одреден инстант (сумарно)"
    (0..2).each { |c| summary.row(sidx).set_format c, title_fmt }
    summary.merge_cells sidx, 0, sidx, 2
    summary.row(sidx).height = 32
    sidx += 1
    summary.row(sidx).push 'id', 'инстант', "бр. неактивни\nтерминали"
    (0..2).each { |c| summary.row(sidx).set_format c, heading_fmt[c] }
    summary.row(sidx).height = 26
    sidx += 1
    
    detailed.column(0).default_format = 
      Spreadsheet::Format.new(horizontal_align: :center, font: font)
    detailed.column(1).default_format = 
      Spreadsheet::Format.new(horizontal_align: :left, font: font)
    detailed.column(0).width = 8
    detailed.column(1).width = 75
    detailed.row(didx)[0] = "Терминали кои не примаат одреден инстант \n(детално)"
    (0..1).each { |c| detailed.row(didx).set_format c, title_fmt }
    detailed.merge_cells didx, 0, didx, 1
    detailed.row(didx).height = 32
    didx += 1

    # loop per instant
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
      summary.row(sidx).push *[ id, Game.find(id).name, inactive_ids.count ]
      (0..2).each { |c| summary.row(sidx).set_format c, sline_fmt[c] }
      sidx += 1
      
      # detailed sheet
      detailed.row(didx)[0] = Game.find(id).name
      detailed.row(didx).set_format 0, 
        Spreadsheet::Format.new(horizontal_align: :center, font: bold_font)
      detailed.merge_cells didx, 0, didx, 1
      didx += 1
      inactive_ids.each do |id|
        detailed.row(didx).push id, Terminal.find(id).name
        didx += 1
      end
    end
  end # create_inactive_sheets

  ##
  # Create sheets for top terminal sales per game
  # 
  def self.create_top_terminals_sheet book, day, opt = {}

    # top 
    top_count = opt[:top_count]
    top_count ||= TOP_COUNT

    # create the worksheet
    tt_sheet   = book.create_worksheet name: 'Топ терминали'

    total_sales = Sale.where("date BETWEEN date(:day, 'start of month') AND :day", day: day).
                    sum(:sales)

    # Formating
    font      =  Spreadsheet::Font.new('Droid Sans') 
    font_sm   =  Spreadsheet::Font.new('Droid Sans', size: 8) 
    bold_font_sm =  Spreadsheet::Font.new('Droid Sans', size: 8, bold: true) 
    bold_font =  Spreadsheet::Font.new('Droid Sans', bold: true) 

    # worksheet default format
    tt_sheet.default_format = Spreadsheet::Format.new font: Spreadsheet::Font.new('Droid Sans')

    # fix some columns widths
    tt_sheet.column(0).width = 10 
    tt_sheet.column(1).width = 75
    tt_sheet.column(2).width = 20
    tt_sheet.column(3).width = 20
    tt_sheet.column(4).width = 15

    # top 2 rows:
    fmt = Spreadsheet::Format.new weight: :bold,
      font: Spreadsheet::Font.new('Droid Sans', bold: true, size: 12)
    tt_sheet.row(0).default_format = fmt
    tt_sheet.row(0).height = 18

    tt_sheet.row(1).default_format = fmt
    tt_sheet.row(1).height = 18
      
    tt_sheet.row(0)[0] = "Период: #{ (day - (day.day - 1)).strftime YMD_FMT } --" +
                         " #{ day.strftime YMD_FMT }"
    tt_sheet.row(1)[0] = "Вкупна уплата: #{ thou_sep(total_sales.to_i) }"
    tt_sheet.merge_cells 0, 0, 0, 4
    tt_sheet.merge_cells 1, 0, 1, 4

    fmt_c = Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :center, font: bold_font)
    fmt_l = Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :left, font: bold_font)
    fmt_r = Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :right, font: bold_font)
    heading_fmt  = [fmt_c, fmt_l, fmt_c, fmt_r, fmt_r]
    tt_sheet.row(3).push *[ "id", "игра/терминал", "цена/град", "промет", "учество %" ]
    tt_sheet.row(3).height = 24
    0.upto(5) { |i| tt_sheet.row(3).set_format i, heading_fmt[i] }

    row = 4
    game_sales = Sale.
        select('s.game_id AS game_id, g.name AS name, g.price AS price, SUM(s.sales) AS total_sales').
        joins('AS s INNER JOIN games AS g ON s.game_id = g.id').
        where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day).
        where('s.sales > 0').
        group('s.game_id').
        order('s.game_id')

    line_fmt_b = [
      Spreadsheet::Format.new(horizontal_align: :center, font: bold_font, vertical_align: :middle),
      Spreadsheet::Format.new(horizontal_align: :left,  font: bold_font, vertical_align: :middle),
      Spreadsheet::Format.new(horizontal_align: :center, font: bold_font, vertical_align: :middle),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',
        vertical_align: :middle, font: bold_font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#0.00%',
        vertical_align: :middle, font: bold_font),
    ]
    line_fmt = [
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :left,  font: font),
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#0.00%',   font: font),
    ]
    game_sales.each do |g|
      r = tt_sheet.row row
      r[0] = g.game_id
      r[1] = g.name
      r[2] = g.price
      r[3] = g.total_sales
      r[4] = g.total_sales*1.0/total_sales
      0.upto(5) {|i| tt_sheet.row(row).set_format i, line_fmt_b[i]}
      r.height = 18
      row += 1

      top_terminals = Sale.
        select('s.terminal_id AS terminal_id, t.name AS name,' +
               ' t.city AS city, SUM(s.sales) AS total_sales').
        joins('AS s INNER JOIN terminals AS t ON s.terminal_id = t.id').
        where("s.date BETWEEN date(:day, 'start of month') AND :day", day: day).
        where('s.game_id' => g.game_id).
        where('s.sales > 0').
        group('s.terminal_id').
        order('total_sales DESC').
        limit(top_count)
      top_terminals.each do |tt|
        r = tt_sheet.row row
        r[0] = tt.terminal_id
        r[1] = tt.name
        r[2] = tt.city
        r[3] = tt.total_sales
        r[4] = tt.total_sales*1.0/g.total_sales
        0.upto(5) {|i| tt_sheet.row(row).set_format i, line_fmt[i]}
        row += 1
      end
    end
  end

  ##
  # Create sheets for sales per city
  # 
  def self.create_sales_per_city_sheet book, day

    # create the worksheet
    sheet   = book.create_worksheet name: 'Продажба по градови'

    # Formating
    font      =  Spreadsheet::Font.new('Droid Sans') 
    font_sm   =  Spreadsheet::Font.new('Droid Sans', size: 8) 
    bold_font_sm =  Spreadsheet::Font.new('Droid Sans', size: 8, bold: true) 
    bold_font =  Spreadsheet::Font.new('Droid Sans', bold: true) 

    # fix some columns widths
    sheet.column(0).width = 25 
    sheet.column(1).width = 15
    sheet.column(2).width = 10
    sheet.column(3).width = 10
    sheet.column(4).width = 12
    sheet.column(5).width = 12
    sheet.column(6).width = 12
    sheet.column(7).width = 10
    sheet.column(8).width = 50


    total_sales = Sale.where("date BETWEEN date(:day, 'start of month') AND :day", day: day).
                    sum(:sales)

    # top 2 rows:
    fmt = Spreadsheet::Format.new weight: :bold,
      font: Spreadsheet::Font.new('Droid Sans', bold: true, size: 12)
    sheet.row(0).default_format = fmt
    sheet.row(0).height = 18

    sheet.row(1).default_format = fmt
    sheet.row(1).height = 18

    sheet.row(0)[0] = "Период: #{ (day - (day.day - 1)).strftime YMD_FMT } --" +
                      " #{ day.strftime YMD_FMT }"
    sheet.row(1)[0] = "Вкупна уплата: #{ thou_sep(total_sales.to_i) }"

    sheet.merge_cells 0, 0, 0, 8
    sheet.merge_cells 1, 0, 1, 8

    # heading
    fmt_c = Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :center, font: bold_font)
    fmt_l = Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :left, font: bold_font)
    fmt_r = Spreadsheet::Format.new(text_wrap: true, vertical_align: :middle, bottom: :thin,
          horizontal_align: :right, font: bold_font)
    heading_fmt  = [fmt_l, fmt_r, fmt_c, fmt_c, fmt_r, fmt_r, fmt_r, fmt_c, fmt_r]
    sheet.row(3).push *%W{град  продажба учество бр.терм. мин. прос. макс. id терминал}
    0.upto(8) { |i| sheet.row(3).set_format i, heading_fmt[i] }
    sheet.row(3).height = 24

    # rows
    row = 4
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

    line_fmt = [
      Spreadsheet::Format.new(horizontal_align: :left, font: font),
      Spreadsheet::Format.new(horizontal_align: :right,  font: font, number_format: '#,###'),
      Spreadsheet::Format.new(horizontal_align: :center, number_format: '#0.00%',   font: font),
      Spreadsheet::Format.new(horizontal_align: :center, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :right, number_format: '#,###',   font: font),
      Spreadsheet::Format.new(horizontal_align: :center, font: font),
      Spreadsheet::Format.new(horizontal_align: :right, font: font),
    ]

    city_sales.each do |s|
      r = sheet.row row
      r[0] = s['city']
      r[1] = s['city_sales']
      r[2] = s['city_sales']*1.0/total_sales
      r[3] = s['term_count']
      r[4] = s['min_term_sales']
      r[5] = s['avg_term_sales']
      r[6] = s['max_term_sales']

      t = term_sales.select {|ts| ts['city'] == s['city'] and ts['sales'] >= s['max_term_sales']}[0]
      r[7] = t['terminal_id']
      r[8] = t['name']

      0.upto(8) {|i| sheet.row(row).set_format i, line_fmt[i]}
      row += 1
    end

  end
end
