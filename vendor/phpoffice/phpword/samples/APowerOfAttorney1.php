<?php
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\Paragraph;
//生成用户委托书
include_once 'Sample_Header.php';

$languageEnGb = new \PhpOffice\PhpWord\Style\Language(\PhpOffice\PhpWord\Style\Language::EN_GB);

$phpWord = new \PhpOffice\PhpWord\PhpWord();
$phpWord->getSettings()->setThemeFontLang($languageEnGb);

$fontStyleName = 'rStyle';
$phpWord->addFontStyle($fontStyleName, array('bold' => true, 'italic' => true, 'size' => 16, 'allCaps' => true, 'doubleStrikethrough' => true));

$paragraphStyleName = 'pStyle';
$phpWord->addParagraphStyle($paragraphStyleName, array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 100));

$phpWord->addTitleStyle(1, array('bold' => true), array('spaceAfter' => 240));

// New portrait section
$section = $phpWord->addSection();

// Simple text
$section->addTitle('Welcome to PhpWord汉字行不行呢', 1);
$section->addText('Hello Worl汉字行不行呢d!');

// $pStyle = new Font();
// $pStyle->setLang()
$section->addText('Ce texte-ci est en français.', array('lang' => \PhpOffice\PhpWord\Style\Language::FR_BE));

// Two text break
$section->addTextBreak(2);

// Define styles
$section->addText('汉字行lllllllllllllll不行呢I am styled by a font style definition.', $fontStyleName);
$section->addText('I am styled by a paragraph style definition.', null, $paragraphStyleName);
$section->addText('I am styled by both font and paragraph style.', $fontStyleName, $paragraphStyleName);

$section->addTextBreak();

// Inline font style
$fontStyle['name'] = 'Times New Roman';
$fontStyle['size'] = 20;

$textrun = $section->addTextRun();
$textrun->addText('I am inline styled ', $fontStyle);
$textrun->addText('with ');
$textrun->addText('color', array('color' => '996699'));
$textrun->addText(', ');
$textrun->addText('bold', array('bold' => true));
$textrun->addText(', ');
$textrun->addText('italic', array('italic' => true));
$textrun->addText(', ');
$section->addTextBreak(1);
$textrun->addText('underline', array('underline' => 'dash'));
$textrun->addText(', ');
$textrun->addText('strikethrough', array('strikethrough' => true));
$textrun->addText(', ');
$textrun->addText('doubleStrikethrough', array('doubleStrikethrough' => true));
$textrun->addText(', ');
$textrun->addText('superScript', array('superScript' => true));
$textrun->addText(', ');
$textrun->addText('subScript', array('subScript' => true));
$textrun->addText(', ');
$textrun->addText('smallCaps', array('smallCaps' => true));
$textrun->addText(', ');
$textrun->addText('allCaps', array('allCaps' => true));
$textrun->addText(', ');
$textrun->addText('fgColor', array('fgColor' => 'yellow'));
$textrun->addText(', ');
$textrun->addText('scale', array('scale' => 200));
$textrun->addText(', ');
$textrun->addText('spacing', array('spacing' => 120));
$textrun->addText(', ');
$textrun->addText('kerning', array('kerning' => 10));
$textrun->addText('. ');

// Link
$section->addLink('https://github.com/PHPOffice/PHPWord', 'PHPWord on GitHub');
$section->addTextBreak();

// Image
$section->addImage('resources/_earth.jpg', array('width'=>18, 'height'=>18));
// Save file
// Write documents
$filename = $_GET["id"];
foreach ($writers as $format => $extension) {
    if (null !== $extension) {
        $targetFile ="../../../../doc/APowerOfAttorney/".$filename.".".$extension;
        $phpWord->save($targetFile, $format);
    }
}
echo "success";
