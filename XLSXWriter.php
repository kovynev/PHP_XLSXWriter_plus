<?php
/*
 * @license MIT License
 * */

if (!class_exists('ZipArchive')) {
    throw new Exception('ZipArchive not found');
}
class XLSXWriter
{
    const DATE_FORMAT = 22;

    //------------------------------------------------------------------
    protected $author = 'Doc Author';

    protected $defaultFontName = 'Calibri';
    protected $defaultFontSize = 11;
    protected $defaultWrapText = false;
    protected $defaultVerticalAlign = 'top';
    protected $defaultHorizontalAlign = 'left';
    protected $defaultStartRow = 0;
    protected $defaultStartCol = 0;

    protected $defaultStyle = array();

    protected $fontsCount = 1; //1 font must be in structure
    protected $fontSize = 8;
    protected $fontColor = '';
    protected $fontStyles = '';
    protected $fontName = '';
    protected $fontId = 0; //font counting from index - 0, means 0,1 - 2 elements

    protected $bordersCount = 1; //1 border must be in structure
    protected $bordersStyle = '';
    protected $bordersColor = '';
    protected $borderId = 0; //borders counting from index - 0, means 0,1 - 2 elements

    protected $fillsCount = 2; //2 fills must be in structure
    protected $fillColor = '';
    protected $fillId = 1; //fill counting from index - 0, means 0,1 - 2 elements

    protected $stylesCount = 1;//1 style must be in structure

    protected $sheets_meta = array();
    protected $sheets = array();
    protected $shared_strings = array();//unique set
    protected $shared_string_count = 0;//count of non-unique references to the unique set
    protected $temp_files = array();

    protected $useSharedStrings = false;

    public function __construct()
    {
    }

    public function setAuthor($author = '')
    {
        $this->author = $author;
    }

    public function setFontName($defaultFontName)
    {
        $this->defaultFontName = $defaultFontName;
    }

    public function setFontSize($defaultFontSize)
    {
        $this->defaultFontSize = $defaultFontSize;
    }

    public function setWrapText($defaultWrapText)
    {
        $this->defaultWrapText = $defaultWrapText;
    }

    public function setVerticalAlign($defaultVerticalAlign)
    {
        $this->defaultVerticalAlign = $defaultVerticalAlign;
    }

    public function setHorizontalAlign($defaultHorizontalAlign)
    {
        $this->defaultHorizontalAlign = $defaultHorizontalAlign;
    }

    private $columnsStyles = [];
    private $rowsStyles = [];
    private $cellsStyles = [];

    private function setStyle($defaultStyle)
    {
        $this->defaultStyle = $defaultStyle;

        foreach ($this->defaultStyle as $styleIndex => $style) {
            if (!array_key_exists('width', $style)) {
                if (!array_key_exists('cells', $style)) {
                    $style['cells'] = [];
                }
                if (!array_key_exists('columns', $style)) {
                    $style['columns'] = [];
                }
                if (!array_key_exists('rows', $style)) {
                    $style['rows'] = [];
                }
                if ($style['cells']) {
                    foreach ($style['cells'] as $cellXlsIndex) {
                        $this->cellsStyles[$cellXlsIndex] = $styleIndex;
                    }
                }
                else if ($style['columns']) {
                    foreach ($style['columns'] as $columnIndex) {
                        $this->columnsStyles[$columnIndex] = $styleIndex;
                    }
                }
                elseif ($style['rows']) {
                    foreach ($style['rows'] as $rowIndex) {
                        $this->rowsStyles[$rowIndex] = $styleIndex;
                    }
                }
            }
        }
    }

    public function setStartRow($defaultStartRow)
    {
        $this->defaultStartRow = ($defaultStartRow > 0) ? ((int)$defaultStartRow - 1) : 0;
    }

    public function setStartCol($defaultStartCol)
    {
        $this->defaultStartCol = ($defaultStartCol > 0) ? ((int)$defaultStartCol - 1) : 0;
    }

    public function __destruct()
    {
        if (!empty($this->temp_files)) {
            foreach ($this->temp_files as $temp_file) {
                @unlink($temp_file);
            }
        }
    }

    protected function tempFilename()
    {
        $filename = tempnam("/tmp", "xlsx_writer_");
        $this->temp_files[] = $filename;
        return $filename;
    }

    /**
     * Put xslx to stdout
     */
    public function writeToStdOut()
    {
        $temp_file = $this->tempFilename();
        self::writeToFile($temp_file);
        readfile($temp_file);
    }

    /**
     * Write to file
     * @return string
     */
    public function writeToString()
    {
        $temp_file = $this->tempFilename();
        self::writeToFile($temp_file);
        $string = file_get_contents($temp_file);
        return $string;
    }

    /**
     * Write spreadsheet to file.
     * @param string $filename
     */
    public function writeToFile($filename)
    {
        @unlink($filename);//if the zip already exists, overwrite it
        $zip = new ZipArchive();
        if (empty($this->sheets_meta)) {
            self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", no worksheets defined.");
            return;
        }
        if (!$zip->open($filename, ZipArchive::CREATE)) {
            self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", unable to create zip.");
            return;
        }

        $zip->addEmptyDir("docProps/");
        $zip->addFromString("docProps/app.xml", self::buildAppXML());
        $zip->addFromString("docProps/core.xml", self::buildCoreXML());

        $zip->addEmptyDir("_rels/");
        $zip->addFromString("_rels/.rels", self::buildRelationshipsXML());

        $zip->addEmptyDir("xl/worksheets/");
        foreach ($this->sheets_meta as $sheet_meta) {
            $zip->addFile($sheet_meta['filename'], "xl/worksheets/" . $sheet_meta['xmlname']);
        }
        if (!empty($this->shared_strings)) {
            $zip->addFile($this->writeSharedStringsXML(), "xl/sharedStrings.xml");  //$zip->addFromString("xl/sharedStrings.xml",     self::buildSharedStringsXML() );
        }
        $zip->addFromString("xl/workbook.xml", self::buildWorkbookXML());
        $zip->addFile($this->writeStylesXML(), "xl/styles.xml");  //$zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
        $zip->addFromString("[Content_Types].xml", self::buildContentTypesXML());

        $zip->addEmptyDir("xl/_rels/");
        $zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelsXML());
        $zip->close();
    }

    private function prepareStyles($styles, $sheetName) {
        for ($i = 0; $i < count($styles); $i++) {
            $styles[$i] += array('sheet' => $sheetName);
        }
        return array_merge((array) $this->defaultStyle, (array) $styles);
    }

    /**
     * @param PDOStatement $query
     * @param callable $rowMapper table columns
     * @param string $sheetName
     * @param array $headersTypes
     * @param array $styles
     * @param array $additionalData
     * @throws Exception
     */
    public function writeSheetFromUnbufferedQuery($query,
                                         $rowMapper,
                                         $sheetName = '',
                                         $headersTypes = [],
                                         $styles = [],
                                         $additionalData = [])
    {
        $styles = $this->prepareStyles($styles, $sheetName);
        $this->setStyle($styles);
        $sheetFilename = $this->tempFilename();

        $sheet_default = 'Sheet' . (count($this->sheets_meta) + 1);
        $sheetName = !empty($sheetName) ? $sheetName : $sheet_default;
        $this->sheets_meta[] = array('filename' => $sheetFilename, 'sheetname' => $sheetName, 'xmlname' => strtolower($sheet_default) . ".xml");

        $headerOffset = empty($headersTypes) ? 0 : $this->defaultStartRow + 1;
        $rowsCount = 0;
        $columnsCount = 0;

        $tabselected = count($this->sheets_meta) == 1 ? 'true' : 'false';//only first sheet is selected
        $cellFormats = array_values($headersTypes);
        $headerRow = array_keys($headersTypes);

        $fd = $this->openSheetFile($sheetFilename);
        $this->writeDocumentHeader($fd, '', $tabselected);
        $this->writeColumnsWidth($fd);

        $this->sheetDataBegins($fd);
        $this->writeAdditionalCells($fd, $additionalData);
        $this->writeDataHeaders($fd, $headerRow);
        $this->writeDataFromQuery($fd, $query, $rowMapper, $headerOffset, $cellFormats, $rowsCount, $columnsCount);

        $this->sheetDataEnds($fd);

        $this->mergeCells($fd, $sheetName);
        $this->writeDocumentBottom($fd);

        $maxCell = self::xlsCell($headerOffset + $rowsCount - 1, $columnsCount - 1);
        $this->finalizeDocument($fd, $maxCell);

        fclose($fd);
    }


    private function openSheetFile($sheetFilename) {
        $fd = fopen($sheetFilename, "w+");
        if ($fd === false) {
            throw new Exception("write failed in " . __CLASS__ . "::" . __FUNCTION__ . ".");
        }
        return $fd;
    }

    private function sheetDataBegins($fd) {
        $this->writeToBuffer($fd, '<sheetData>');
    }

    private function sheetDataEnds($fd) {
        $this->writeToBuffer($fd, '</sheetData>');
    }

    private function writeDocumentHeader($fd, $maxCell, $tabselected) {
        $this->writeToBuffer($fd, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . PHP_EOL);
        $this->writeToBuffer($fd, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                            mc:Ignorable="x14ac"
                            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'. PHP_EOL);
        $this->writeToBuffer($fd, '<sheetPr filterMode="false">'. PHP_EOL);
        $this->writeToBuffer($fd, '<pageSetUpPr fitToPage="false"/>'. PHP_EOL);
        $this->writeToBuffer($fd, '</sheetPr>'. PHP_EOL);
        $this->writeToBuffer($fd, '<dimension ref="A1:' . $maxCell . '                         "/>'. PHP_EOL);
        $this->writeToBuffer($fd, '<sheetViews>'. PHP_EOL);
        $this->writeToBuffer($fd, '<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="' . $tabselected . '" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">'. PHP_EOL);
        $this->writeToBuffer($fd, '<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>'. PHP_EOL);
        $this->writeToBuffer($fd, '</sheetView>'. PHP_EOL);
        $this->writeToBuffer($fd, '</sheetViews>'. PHP_EOL);
    }

    private function correctDocumentDimension($fd, $maxCell) {
        fseek($fd, 0);
        while (!feof($fd)) {
            $lastPos = ftell($fd);
            $line = fgets($fd);
            if ($line) {
                if (stripos($line, '<dimension ref="A1:') !== false) {
                    $writePos = $lastPos + 19;
                    fseek($fd, $writePos);
                    fwrite($fd, $maxCell);
                    fwrite($fd, '"/>');
                    if (($enclosedTagPos = strripos($line, '"/>')) !== false) {
                        fseek($fd, $lastPos + $enclosedTagPos);
                        fwrite($fd, '   ');
                    }
                    break;
                }
            }
        }
        fseek($fd, 0, SEEK_END);
    }

    private function writeColumnsWidth($fd) {
        $this->writeToBuffer($fd, '<cols>');
        //fetch all columns with custom width
        $customWidthColumns = [];
        foreach ($this->defaultStyle as $style) {
            if (array_key_exists('width', $style) && array_key_exists('columns', $style)) {
                foreach ($style['columns'] as $columnNumber) {
                    $customWidthColumns[$columnNumber] = $style['width'];
                }
                $customWidthColumns += $style['columns'];
            }
        }
        ksort($customWidthColumns);
        $i = 1;
        foreach ($customWidthColumns as $columnNumber => $width) {
            $this->writeToBuffer($fd, sprintf('<col min="%d" max="%d" customWidth="true" width="%s" style="0" />'. PHP_EOL, $i, $i, $width));
            $i++;
        }
        $this->writeToBuffer($fd, '</cols>'. PHP_EOL);
    }

    /**
     * @param $fd
     * @param array $headerRow
     */
    private function writeDataHeaders($fd, $headerRow) {
        if (!empty($headerRow)) {
            $this->writeToBuffer($fd, '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($this->defaultStartRow + 1) . '">'. PHP_EOL);
            foreach ($headerRow as $k => $v) {
                $this->writeCell($fd, $this->defaultStartRow + 0, $this->defaultStartCol + $k, $v);
            }
            $this->writeToBuffer($fd, '</row>'. PHP_EOL);
        }
    }

    /**
     * @param $fd
     * @param PDOStatement $query
     * @param callable $rowMapper
     * @param $headerOffset
     * @param $cellFormats
     * @param $rowsCount
     * @param $columnsCount
     */
    private function writeDataFromQuery($fd, $query, $rowMapper, $headerOffset, $cellFormats, &$rowsCount, &$columnsCount) {

        $i = 0;
        $columnsCount = null;

        while ($row = $query->fetch(PDO::FETCH_ASSOC)) {
            if ($row) {
                $this->writeToBuffer($fd, '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.($i + $headerOffset + 1).'">'. PHP_EOL);
                $row = $rowMapper($row);
                if ($columnsCount == null) {
                    $columnsCount = count($row);
                }

                $columnOffset = 0;
                foreach($row as $key => $value) {
                    $rowNumber = $i + $headerOffset;
                    $columnNumber = $this->defaultStartCol + $columnOffset;
                    $cellType = $cellFormats[$columnNumber];

                    $cell = self::xlsCell($rowNumber, $columnNumber);
                    $s = 0;
                    if (isset($this->cellsStyles[$cell])){
                        $s = $this->cellsStyles[$cell] + 1;
                    }
                    elseif (isset($this->rowsStyles[$rowNumber])) {
                        $s = $this->rowsStyles[$rowNumber] + 1;
                    }
                    elseif (isset($this->columnsStyles[$columnNumber])) {
                        $s = $this->columnsStyles[$columnNumber] + 1;
                    }
                    if (is_numeric($value)) {
                        $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="n"><v>' . ($value * 1) . '</v></c>'. PHP_EOL); //int, float, etc
                    } else if ($cellType == 'date') {
                        $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="n"><v>' . (int) (self::convertDateTime($value)) . '</v></c>'. PHP_EOL);
                    } else if ($cellType == 'datetime') {
                        $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '"><v>' . self::convertDateTime($value) . '</v></c>'. PHP_EOL);
                    } else if ($value == '') {
                        $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '"/>'. PHP_EOL);
                    } else if ($value{0} == '=') {
                        $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="s"><f>' . self::xmlspecialchars($value) . '</f></c>'. PHP_EOL);
                    } else if ($value !== '') {
                        $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="inlineStr"><is><t>' . self::xmlspecialchars($value) . '</t></is></c>'. PHP_EOL);
                    }

                    $columnOffset++;
                }

                $this->writeToBuffer($fd, '</row>'. PHP_EOL);

                $rowsCount++;
            }
            $i++;
        }
    }

    private $writeOperationsCount = 0;
    private $writeBuffer = '';

    /**
     * It is faster to concat strings and write it to file later.
     * This method puts string to file once at 1000 calls
     * @param $fd
     * @param $string
     */
    private function writeToBuffer($fd, $string) {
        $this->writeBuffer .= $string;
        $this->writeOperationsCount++;
        if($this->writeOperationsCount > 1000) {
            $this->dumpBufferToFile($fd);
        }
    }

    /**
     * Push data to file if we have a tail
     * @param $fd
     */
    private function dumpBufferToFile($fd) {
        if ($this->writeBuffer) {
            fwrite($fd, $this->writeBuffer);
            $this->writeBuffer = '';
            $this->writeOperationsCount = 0;
        }
    }

    /**
     * @param $fd
     * @param array $cellsData
     */
    private function writeAdditionalCells($fd, $cellsData) {
        $additionalDataMapped = [];
        foreach ($cellsData as $xlsCell => $row) {
            $zeroBasedCoordinates = self::cellCoordinates($xlsCell);
            list($i, $k) = $zeroBasedCoordinates;
            $additionalDataMapped[$i][$k] = $row;
        }
        ksort($additionalDataMapped);
        foreach ($additionalDataMapped as $i => $row) {
            $this->writeToBuffer($fd, '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($i + 1) . '">'. PHP_EOL);
            foreach ($row as $k => $v) {
                $cellValue = $v;
                $cellType = '';
                if (is_array($v)) {
                    list($cellValue, $cellType) = $v;
                }
                $this->writeCell($fd, $i, $k, $cellValue, $cellType);
            }
            $this->writeToBuffer($fd, '</row>'. PHP_EOL);
        }
    }

    /**
     * @param $fd
     * @param string $sheetName
     */
    private function mergeCells($fd, $sheetName) {
        $sheet = $this->sheets[$sheetName];
        if (!empty($sheet->merge_cells)) {
            $this->writeToBuffer($fd, '<mergeCells>');
            foreach ($sheet->merge_cells as $range) {
                $this->writeToBuffer($fd, '<mergeCell ref="' . $range . '"/>');
            }
            $this->writeToBuffer($fd, '</mergeCells>');
        }
    }

    private function writeDocumentBottom($fd) {
        $this->writeToBuffer($fd, '<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>'. PHP_EOL);
        $this->writeToBuffer($fd, '<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>'. PHP_EOL);
        $this->writeToBuffer($fd, '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>'. PHP_EOL);
        $this->writeToBuffer($fd, '<headerFooter differentFirst="false" differentOddEven="false">'. PHP_EOL);
        $this->writeToBuffer($fd, '<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>'. PHP_EOL);
        $this->writeToBuffer($fd, '<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>'. PHP_EOL);
        $this->writeToBuffer($fd, '</headerFooter>'. PHP_EOL);
        $this->writeToBuffer($fd, '</worksheet>'. PHP_EOL);
    }

    private function finalizeDocument($fd, $maxCell)
    {
        $this->dumpBufferToFile($fd);
        $this->correctDocumentDimension($fd, $maxCell);
    }

    /**
     * Merge cells
     * @param string $sheetName
     * @param int $startCellRow
     * @param int $startCellColumn
     * @param int $endCellRow
     * @param int $endCellColumn
     */
    public function markMergedCell($sheetName, $startCellRow, $startCellColumn, $endCellRow, $endCellColumn)
    {
        $sheet = &$this->sheets[$sheetName];
        $startCell = self::xlsCell($startCellRow, $startCellColumn);
        $endCell = self::xlsCell($endCellRow, $endCellColumn);
        $sheet->merge_cells[] = $startCell . ":" . $endCell;
    }

    /**
     * @param bool $value
     */
    public function useSharedStrings($value) {
        $this->useSharedStrings = $value;
    }

    /**
     * @param $fd
     * @param int $rowNumber
     * @param int $columnNumber
     * @param $value
     * @param string $cellType
     * @internal param string $sheetName
     */
    protected function writeCell($fd, $rowNumber, $columnNumber, $value, $cellType = '')
    {
        $cell = self::xlsCell($rowNumber, $columnNumber);
        $s = 0;
        if ($this->defaultStyle) {
            if (isset($this->cellsStyles[$cell])){
                $s = $this->cellsStyles[$cell] + 1;
            }
            elseif (isset($this->rowsStyles[$rowNumber])) {
                $s = $this->rowsStyles[$rowNumber] + 1;
            }
            elseif (isset($this->columnsStyles[$columnNumber])) {
                $s = $this->columnsStyles[$columnNumber] + 1;
            }
        }
        if (is_numeric($value)) {
            $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="n"><v>' . ($value * 1) . '</v></c>'. PHP_EOL); //int, float, etc
        } else if ($cellType == 'date') {
            $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="n"><v>' . intval(self::convertDateTime($value)) . '</v></c>'. PHP_EOL);
        } else if ($cellType == 'datetime') {
            $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '"><v>' . self::convertDateTime($value) . '</v></c>'. PHP_EOL);
        } else if ($value == '') {
            $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '"/>'. PHP_EOL);
        } else if ($value{0} == '=') {
            $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="s"><f>' . self::xmlspecialchars($value) . '</f></c>'. PHP_EOL);
        } else if ($value !== '') {
            if ($this->useSharedStrings) {
                $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="s"><v>' . self::xmlspecialchars($this->setSharedString($value)) . '</v></c>'. PHP_EOL);
            }
            else {
                $this->writeToBuffer($fd, '<c r="' . $cell . '" s="' . $s . '" t="inlineStr">' .
                    '<is>'.
                      '<t>'. self::xmlspecialchars($value) . '</t>'.
                    '</is>'.
                    '</c>'. PHP_EOL);
            }
        }
    }

    /**
     * @return string
     */
    protected function writeStylesXML()
    {
        $tempfile = $this->tempFilename();
        $fd = fopen($tempfile, "w+");
        if ($fd === false) {
            self::log("write failed in " . __CLASS__ . "::" . __FUNCTION__ . ".");
            return;
        }
        fwrite($fd, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'. PHP_EOL);
        fwrite($fd, '<styleSheet xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'. PHP_EOL);
        if ($this->defaultStyle) {
            foreach ($this->defaultStyle as $style) {
                if (isset($style['sheet'])) {
                    if (isset($style['font'])) $this->fontsCount++;
                }
            }
        }
        fwrite($fd, '<fonts x14ac:knownFonts="1" count="' . $this->fontsCount . '">'. PHP_EOL);
        fwrite($fd, '	<font>'. PHP_EOL);
        fwrite($fd, '		<sz val="' . $this->defaultFontSize . '"/>'. PHP_EOL);
        fwrite($fd, '		<color theme="1"/>');
        fwrite($fd, '		<name val="' . $this->defaultFontName . '"/>'. PHP_EOL);
        fwrite($fd, '		<family val="2"/>'. PHP_EOL);
        if ($this->defaultFontName == 'MS Sans Serif') {
            fwrite($fd, '		<charset val="204"/>'. PHP_EOL);
        } else if ($this->defaultFontName == 'Calibri') {
            fwrite($fd, '		<scheme val="minor"/>'. PHP_EOL);
        } else {
            fwrite($fd, '		<charset val="204"/>'. PHP_EOL);
        }
        fwrite($fd, '	</font>');
        if ($this->defaultStyle) {
            foreach ($this->defaultStyle as $style) {
                if (isset($style['sheet'])) {
                    if (isset($style['font'])) {
                        if (isset($style['font']['name']) && !empty($style['font']['name'])) $this->fontName = $style['font']['name'];
                        if (isset($style['font']['size']) && !empty($style['font']['size'])) $this->fontSize = $style['font']['size'];
                        if (isset($style['font']['color']) && !empty($style['font']['color'])) $this->fontColor = $style['font']['color'];
                        if (isset($style['font']['bold']) && !empty($style['font']['bold'])) $this->fontStyles .= '<b/>';
                        if (isset($style['font']['italic']) && !empty($style['font']['italic'])) $this->fontStyles .= '<i/>';
                        if (isset($style['font']['underline']) && !empty($style['font']['underline'])) $this->fontStyles .= '<u/>';

                        fwrite($fd, '	<font>');
                        if ($this->fontStyles) fwrite($fd, '		' . $this->fontStyles);
                        fwrite($fd, '		<sz val="' . $this->fontSize . '"/>');
                        if ($this->fontColor) {
                            fwrite($fd, '		<color rgb="FF' . $this->fontColor . '"/>');
                        } else {
                            fwrite($fd, '		<color theme="1"/>');
                        }
                        if ($this->fontName) {
                            fwrite($fd, '		<name val="' . $this->fontName . '"/>');
                        }
                        fwrite($fd, '		<family val="2"/>');
                        if ($this->fontName == 'MS Sans Serif') {
                            fwrite($fd, '		<charset val="204"/>');
                        } else if ($this->fontName == 'Calibri') {
                            fwrite($fd, '		<scheme val="minor"/>');
                        } else {
                            fwrite($fd, '		<charset val="204"/>');
                        }
                        fwrite($fd, '	</font>');
                    }
                    $this->fontStyles = '';
                }
            }
        }
        fwrite($fd, '</fonts>');
        if ($this->defaultStyle) {
            foreach ($this->defaultStyle as $style) {
                if (isset($style['sheet'])) {
                    if (isset($style['fill'])) $this->fillsCount++;
                }
            }
        }
        fwrite($fd, '<fills count="' . $this->fillsCount . '">');
        fwrite($fd, '	<fill><patternFill patternType="none"/></fill>');
        fwrite($fd, '	<fill><patternFill patternType="gray125"/></fill>');
        if ($this->defaultStyle) {
            foreach ($this->defaultStyle as $style) {
                if (isset($style['sheet'])) {
                    if (isset($style['fill'])) {
                        if (isset($style['fill']['color'])) $this->fillColor = $style['fill']['color'];
                        fwrite($fd, '	<fill>');
                        fwrite($fd, '		<patternFill patternType="solid">');
                        fwrite($fd, '			<fgColor rgb="FF' . $this->fillColor . '"/>');
                        fwrite($fd, '			<bgColor indexed="64"/>');
                        fwrite($fd, '		</patternFill>');
                        fwrite($fd, '	</fill>');
                    }
                }
            }
        }
        fwrite($fd, '</fills>');
        if ($this->defaultStyle) {
            foreach ($this->defaultStyle as $style) {
                if (isset($style['sheet'])) {
                    if (isset($style['border'])) $this->bordersCount++;
                }
            }
        }
        fwrite($fd, '<borders count="' . $this->bordersCount . '">');
        fwrite($fd, '	<border>');
        fwrite($fd, '		<left/><right/><top/><bottom/><diagonal/>');
        fwrite($fd, '	</border>');
        if ($this->defaultStyle) {
            foreach ($this->defaultStyle as $style) {
                if (isset($style['sheet'])) {
                    if (isset($style['border'])) {
                        if (isset($style['border']['style'])) $this->bordersStyle = ' style="' . $style['border']['style'] . '"';
                        if (isset($style['border']['color'])) $this->bordersColor = '<color rgb="FF' . $style['border']['color'] . '"/>';
                        fwrite($fd, '	<border>');
                        fwrite($fd, '		<left' . $this->bordersStyle . '>' . $this->bordersColor . '</left>');
                        fwrite($fd, '		<right' . $this->bordersStyle . '>' . $this->bordersColor . '</right>');
                        fwrite($fd, '		<top' . $this->bordersStyle . '>' . $this->bordersColor . '</top>');
                        fwrite($fd, '		<bottom' . $this->bordersStyle . '>' . $this->bordersColor . '</bottom>');
                        fwrite($fd, '		<diagonal/>');
                        fwrite($fd, '	</border>');
                    }
                }
            }
        }
        fwrite($fd, '</borders>');
        fwrite($fd, '<cellStyleXfs count="1">');
        fwrite($fd, '<xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        fwrite($fd, '</cellStyleXfs>');
        $this->stylesCount += count($this->defaultStyle);
        fwrite($fd, '<cellXfs count="' . $this->stylesCount . '">');
        $this->defaultWrapText = ($this->defaultWrapText) ? '1' : '0';
        fwrite($fd, '<xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"><alignment wrapText="' . $this->defaultWrapText . '" vertical="' . $this->defaultVerticalAlign . '" horizontal="' . $this->defaultHorizontalAlign . '"/></xf>');
        if ($this->defaultStyle) {
            foreach ($this->defaultStyle as $style) {
                if (isset($style['sheet'])) {
                    if (isset($style['font'])) {
                        $font_Id = $this->fontId += 1;
                    } else {
                        $font_Id = 0;
                    }
                    if (isset($style['fill'])) {
                        $fill_Id = $this->fillId += 1;
                    } else {
                        $fill_Id = 0;
                    }
                    if (isset($style['border'])) {
                        $border_Id = $this->borderId += 1;
                    } else {
                        $border_Id = 0;
                    }
                    if (isset($style['wrapText'])) {
                        $wrapText = ($style['wrapText']) ? '1' : '0';
                    } else {
                        $wrapText = $this->defaultWrapText;
                    }

                    $format_Id = (isset($style['format'])) ? $style['format'] : '0';

                    if (isset($style['verticalAlign'])) {
                        $verticalAlign = $style['verticalAlign'];
                    } else {
                        $verticalAlign = $this->defaultVerticalAlign;
                    }
                    if (isset($style['horizontalAlign'])) {
                        $horizontalAlign = $style['horizontalAlign'];
                    } else {
                        $horizontalAlign = $this->defaultHorizontalAlign;
                    }
                    fwrite($fd, '<xf borderId="' . $border_Id . '" fillId="' . $fill_Id . '" fontId="' . $font_Id . '" numFmtId="' . $format_Id . '" xfId="0" applyFill="1">');
                    fwrite($fd, '<alignment wrapText="' . $wrapText . '" vertical="' . $verticalAlign . '" horizontal="' . $horizontalAlign . '"/>');
                    fwrite($fd, '</xf>');
                }
            }
        }
        fwrite($fd, '</cellXfs>');
        fwrite($fd, '<cellStyles count="1">');
        fwrite($fd, '<cellStyle xfId="0" builtinId="0" name="Normal"/>');
        fwrite($fd, '</cellStyles>');
        fwrite($fd, '<dxfs count="0"/>');
        fwrite($fd, '<tableStyles count="0" defaultPivotStyle="PivotStyleMedium9" defaultTableStyle="TableStyleMedium2"/>');
        fwrite($fd, '<extLst>');
        fwrite($fd, '<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}">');
        fwrite($fd, '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>');
        fwrite($fd, '</ext>');
        fwrite($fd, '</extLst>');
        fwrite($fd, '</styleSheet>');
        fclose($fd);
        return $tempfile;
    }

    private $shared_string_counter = 0;

    protected function setSharedString($v)
    {
        if (isset($this->shared_strings[$v])) {
            $string_value = $this->shared_strings[$v];
        } else {
            $string_value = $this->shared_string_counter;
            $this->shared_strings[$v] = $string_value;
            $this->shared_string_counter++;
        }
        $this->shared_string_count++;//non-unique count
        return $string_value;
    }

    protected function writeSharedStringsXML()
    {
        $tempfile = $this->tempFilename();
        $fd = fopen($tempfile, "w+");
        if ($fd === false) {
            self::log("write failed in " . __CLASS__ . "::" . __FUNCTION__ . ".");
            return;
        }

        fwrite($fd, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        fwrite($fd, '<sst count="' . ($this->shared_string_count) . '" uniqueCount="' . count($this->shared_strings) . '" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
        foreach ($this->shared_strings as $s => $c) {
            fwrite($fd, '<si><t>' . self::xmlspecialchars($s) . '</t></si>');
        }
        fwrite($fd, '</sst>');
        fclose($fd);
        return $tempfile;
    }

    protected function buildAppXML()
    {
        $app_xml = "";
        $app_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $app_xml .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime></Properties>';
        return $app_xml;
    }

    protected function buildCoreXML()
    {
        $core_xml = "";
        $core_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $core_xml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
        $core_xml .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . date("Y-m-d\TH:i:s.00\Z") . '</dcterms:created>';//$date_time = '2013-07-25T15:54:37.00Z';
        $core_xml .= '<dc:creator>' . self::xmlspecialchars($this->author) . '</dc:creator>';
        $core_xml .= '<cp:revision>0</cp:revision>';
        $core_xml .= '</cp:coreProperties>';
        return $core_xml;
    }

    protected function buildRelationshipsXML()
    {
        $rels_xml = "";
        $rels_xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $rels_xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $rels_xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
        $rels_xml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
        $rels_xml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
        $rels_xml .= "\n";
        $rels_xml .= '</Relationships>';
        return $rels_xml;
    }

    protected function buildWorkbookXML()
    {
        $workbook_xml = "";
        $workbook_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $workbook_xml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $workbook_xml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
        $workbook_xml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $workbook_xml .= '<sheets>';
        foreach ($this->sheets_meta as $i => $sheet_meta) {
            $workbook_xml .= '<sheet name="' . self::xmlspecialchars($sheet_meta['sheetname']) . '" sheetId="' . ($i + 1) . '" state="visible" r:id="rId' . ($i + 2) . '"/>';
        }
        $workbook_xml .= '</sheets>';
        $workbook_xml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
        return $workbook_xml;
    }

    protected function buildWorkbookRelsXML()
    {
        $wkbkrels_xml = "";
        $wkbkrels_xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $wkbkrels_xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $wkbkrels_xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        foreach ($this->sheets_meta as $i => $sheet_meta) {
            $wkbkrels_xml .= '<Relationship Id="rId' . ($i + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' . ($sheet_meta['xmlname']) . '"/>';
        }
        if (!empty($this->shared_strings)) {
            $wkbkrels_xml .= '<Relationship Id="rId' . (count($this->sheets_meta) + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
        }
        $wkbkrels_xml .= "\n";
        $wkbkrels_xml .= '</Relationships>';
        return $wkbkrels_xml;
    }

    protected function buildContentTypesXML()
    {
        $content_types_xml = "";
        $content_types_xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $content_types_xml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $content_types_xml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $content_types_xml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        foreach ($this->sheets_meta as $i => $sheet_meta) {
            $content_types_xml .= '<Override PartName="/xl/worksheets/' . ($sheet_meta['xmlname']) . '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        if (!empty($this->shared_strings)) {
            $content_types_xml .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
        }
        $content_types_xml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        $content_types_xml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $content_types_xml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $content_types_xml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $content_types_xml .= "\n";
        $content_types_xml .= '</Types>';

        return $content_types_xml;
    }

    /**
     * Cache column names (A, B, C, D) to improve method performance
     * @var array
     */
    private static $cachedXlsColumnNames = [];
    /**
     * @param int $rowNumber zero based
     * @param int $columnNumber zero based
     * @return string Cell label/coordinates, ex: A1, C3, AA42
     */
    public static function xlsCell($rowNumber, $columnNumber)
    {
        if (!array_key_exists($columnNumber, self::$cachedXlsColumnNames)) {
            $n = $columnNumber;
            for ($xlsColumnName = ""; $n >= 0; $n = intval($n / 26) - 1) {
                $xlsColumnName = chr($n % 26 + 0x41) . $xlsColumnName;
            }
            self::$cachedXlsColumnNames[$columnNumber] = $xlsColumnName;
        }
        else {
            $xlsColumnName = self::$cachedXlsColumnNames[$columnNumber];
        }


        return $xlsColumnName . ($rowNumber + 1);
    }

    /**
     * @param string $xlsCell
     * @return array [rowIndex, columnIndex] zero based
     */
    public static function cellCoordinates($xlsCell) {
        preg_match('/([a-zA-Z]{1,})([0-9]{1,})/', $xlsCell, $matches);
        $letters = $matches[1];
        $rowNumber = $matches[2];

        $columnNumber = 0;
        $order = 0;
        for($i = strlen($letters) - 1; $i >= 0; $i--) {
            $letterByte = ord($letters[$i]) - 0x40;
            $columnNumber += $letterByte * pow(26, $order);
            $order++;
        }
        $columnNumber--;
        $rowNumber--;

        return [$rowNumber, $columnNumber];
    }

    //------------------------------------------------------------------
    public static function log($string)
    {
        file_put_contents("php://stderr", date("Y-m-d H:i:s:") . rtrim(is_array($string) ? json_encode($string) : $string) . "\n");
    }

    //------------------------------------------------------------------
    public static function xmlspecialchars($val)
    {
        return str_replace("'", "&#39;", htmlspecialchars($val));
    }

    //------------------------------------------------------------------
    public static function array_first_key(array $arr)
    {
        reset($arr);
        $first_key = key($arr);
        return $first_key;
    }

    //------------------------------------------------------------------
    public static function convertDateTime($date_time) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
    {
        $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;

        //2016-03-02 12:43:00
        if (strlen($date_time) >= 9) {
            $year = (int)($date_time[0] . $date_time[1] . $date_time[2] . $date_time[3]);
            $month = (int)($date_time[5] . $date_time[6]);
            $day = (int)($date_time[8] . $date_time[9]);
        }

        if (strlen($date_time) == 19) {
            $hour = (int)($date_time[11] . $date_time[12]);
            $min = (int)($date_time[14] . $date_time[15]);
            $sec = (int)($date_time[17] . $date_time[18]);

            $seconds = ($hour * 3600 + $min * 60 + $sec) / 86400;
        }

        //using 1900 as epoch, not 1904, ignoring 1904 special case
        # Special cases for Excel.
        if ($year == 1899 && $month == 12 && $day == 31) return $seconds;    # Excel 1900 epoch
        if ($year == 1900 && $month == 01 && $day == 00) return $seconds;    # Excel 1900 epoch
        if ($year == 1900 && $month == 02 && $day == 29) return 60 + $seconds;    # Excel false leapday

        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch = 1900;
        $offset = 0;
        $norm = 300;
        $range = $year - $epoch;

        # Set month days and check for leap year.
        $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100))) ? 1 : 0;

        $februaryDays = ($leap ? 29 : 28);

        $daysInMonth = 31;
        if ($month == 2) {
            $daysInMonth = $februaryDays;
        } elseif ($month == 4 || $month == 6 || $month == 9 || $month == 11) {
            $daysInMonth = 30;
        }
        # Some boundary checks
        if ($year < $epoch || $year > 9999) return 0;
        if ($month < 1 || $month > 12) return 0;
        if ($day < 1 || $day > $daysInMonth) return 0;

        # Accumulate the number of days since the epoch.
        $days = $day;    # Add days for current month
        $lastMonth = $month - 1;
        switch ($lastMonth) {
            case 1: {
                $days += 31;
            }
                break;
            case 2: {
                $days += 31 + $februaryDays;
            }
                break;
            case 3: {
                $days += $februaryDays + 62;
            }
                break;
            case 4: {
                $days += $februaryDays + 92;
            }
                break;
            case 5: {
                $days += $februaryDays + 123;
            }
                break;
            case 6: {
                $days += $februaryDays + 153;
            }
                break;
            case 7: {
                $days += $februaryDays + 184;
            }
                break;
            case 8: {
                $days += $februaryDays + 215;
            }
                break;
            case 9: {
                $days += $februaryDays + 245;
            }
                break;
            case 10: {
                $days += $februaryDays + 276;
            }
                break;
            case 11: {
                $days += $februaryDays + 306;
            }
                break;
            case 12: {
                $days += $februaryDays + 337;
            }
                break;
        }

        $days += $range * 365;                      # Add days for past years
        $days += (int)(($range) / 4);             # Add leapdays
        $days -= (int)(($range + $offset) / 100); # Subtract 100 year leapdays
        $days += (int)(($range + $offset + $norm) / 400);  # Add 400 year leapdays
        $days -= $leap;                                      # Already counted above

        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }
    //------------------------------------------------------------------
}