<?php
#
#       Spreadsheet.php - a library for generating spreadsheets
#          in Excel's XML format
#
#       (c) 2014 Kathryn Lybarger. CC-BY-SA
#

class Spreadsheet { # actually a workbook, but who says that?

	public $styles; # array of type Style
	public $sheets; # array of type Sheet
	public $activeSheet; # integer

	public $default_style;
	public $title_style;

	function __construct( ) {
		$this->default_style = new Style("Default", "Normal");
		$this->title_style = new Style("Title", "Title", 
			array(
				'FontColor' => '#ffffff',
				'FontBold' => true,
				'BackgroundColor' => '#00ff00',
			)
		);
		$this->link_style = new Style("Hyperlink", "Hyperlink",
			array(
				'FontColor' => '#0000D4',
				'Underline' => 'Single'
			)
		);
		$this->styles = array( 
			$this->default_style,
			$this->title_style,
			$this->link_style
		);
		$this->sheets = array( new Sheet("Sheet1") ) ;
		$this->activeSheet = 0;
	}

	
	public function asXML() {
		$str = '<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
  <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
    <ActiveSheet>' . $this->activeSheet . '</ActiveSheet>
  </ExcelWorkbook>
  <Styles>
';
		foreach ($this->styles as $style) {
			$fontbold = ($style->FontBold)?' ss:Bold="1"':"";
			$bgcolor = ($style->BackgroundColor) ?
				' ss:Color="' . $style->BackgroundColor . '" ss:Pattern="Solid"'
				: "";
			$underline = ($style->Underline)?' ss:Underline="' . 
				$style->Underline . '"':"";
			$numberformat = ($style->NumberFormat)?' ss:Format="' . 
				$style->NumberFormat . '"':"";
			$str .= '    <Style ss:ID="' . $style->ID . '" ss:Name="' .
				$style->name . '">
      <Alignment ss:Vertical="Bottom"/>
      <Borders/>
      <Font ss:FontName="' . $style->FontName .'" x:Family="' . 
		$style->FontFamily .'" ss:Size="' . $style->FontSize . 
		'" ss:Color="' . $style->FontColor . '"' . $fontbold . $underline . '/>
      <Interior' . $bgcolor . '/>
      <NumberFormat' . $numberformat . '/>
      <Protection/>
    </Style>
';
		}
		$str .= '  </Styles>
';
		foreach ($this->sheets as $sheet) {
			$str .= '  <Worksheet ss:Name="' . $sheet->name . '">
    <Table>
';

			foreach ($sheet->columns as $column) {
				if ($column->style != "Default") {
					$colstyle = ' ss:StyleID="' . $column->style . '"';
				} else {
					$colstyle = "";
				}
				$str .= '      <Column ss:AutoFitWidth="0" ss:Width="' . 
					$column->width . '"' . $colstyle . '/>
';
			}
			foreach ($sheet->rows as $row) {
				$str .= '      <Row';
				if ($row->style != 'Default') {
					$str .= ' ss:StyleID="' . $row->style . '"';
				}
				$str .= '>
';
				foreach ($row->cells as $cell) {
					$str .= '        <Cell';
					if ($cell->style != 'Default') {
						$str .= ' ss:StyleID="' . $cell->style . '"';
					}
					if ($cell->href) {
						$str .= ' ss:HRef="' . $cell->href . '"';
					}
					$str .= '>
';
					$str .= '          <Data ss:Type="' . $cell->data->type . '">';
					$str .= htmlspecialchars($cell->data->content) . '</Data>
';
					$str .= '        </Cell>
';
				}
				$str .= '      </Row>
';
			}
			$str .= '    </Table>
';
			$str .= '  </Worksheet>
';
		}
		$str .= '</Workbook>
';

		return $str;
	}

}

class Style {
	public $ID;
	public $name;
	public $FontName;
	public $FontFamily;
	public $FontSize;
	public $FontColor;
	public $FontBold;
	public $BackgroundColor;
	public $Underline;
	public $NumberFormat;

	function __construct( $ID, $name, $op = array() ) {
		$this->ID = $ID;
		$this->name = $name;
		$this->FontName = isset($op['FontName'])?$op['FontName']:"Calibri";
		$this->FontFamily = isset($op['FontFamily'])?$op['FontFamily']:"Swiss";
		$this->FontSize = isset($op['FontSize'])?$op['FontSize']:"11";
		$this->FontColor = isset($op['FontColor'])?$op['FontColor']:"#000000";
		$this->FontBold = isset($op['FontBold'])?$op['FontBold']:false;
		$this->BackgroundColor = isset($op['BackgroundColor'])?$op['BackgroundColor']:null;
		$this->Underline = isset($op['Underline'])?$op['Underline']:null;
		$this->NumberFormat = isset($op['NumberFormat'])?$op['NumberFormat']:null;
	}
}

class Sheet { # a two-by-two array of data
	public $name;
	public $columns;
	public $rows;

	function __construct( $name, $op = array() ) {
		$this->name = $name;
		$this->columns = array();
		$this->rows = array();
	}

	public function appendColumn($column) {
		$this->columns[] = $column;
	}

	public function appendRow($row) {
		$this->rows[] = $row;
	}

	function addTitleRow( $data, $op = array() ) {
		$row = new Row();
		$row->populate( $data );
		$row->style = "Title";
		$this->rows = array($row);
	}
	
	public function deleteAllRows() {
		$this->rows = array();
	}
}

class Column { # contains no data, mainly style
	public $width;
	public $style;

	function __construct( $op = array() ) {
		$this->width = (isset($op['width']))?$op['width']:null;
		$this->style = (isset($op['style']))?$op['style']:"Default";
	}
}

class Row {  
	public $cells;
	public $style;

	function __construct( $op = array() ) {
		$this->style = "Default";
	}

	function populate( $data ) {
		$this->cells = array();
		foreach ($data as $datum) {
			$this->cells[] = new Cell($datum);
		}
	}
}

class Cell {
	public $href;
	public $style;
	public $data;

	function __construct( $data, $op = array() ) {
		$this->style = "Default";
		$this->data = new Data($data);
	}
}

class Data {
	public $type;
	public $content;

	function __construct( $data, $op = array() ) {
		$this->content = $data;
		$this->type = "String";
	}
}

