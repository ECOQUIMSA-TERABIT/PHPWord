<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2018 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Element;

use PhpOffice\PhpWord\Writer\Word2007\Style\Image as ImageStyleWriter;

/**
 * OLEObject element writer
 *
 * @since 0.10.0
 */
class OLEObject extends AbstractElement
{
    /**
     * Write object element.
     */
    public function write()
    {
        $xmlWriter = $this->getXmlWriter();
        $element = $this->getElement();
        if (!$element instanceof \PhpOffice\PhpWord\Element\OLEObject) {
            return;
        }

        $rIdObject = $element->getRelationId() + ($element->isInSection() ? 6 : 0);
        $rIdImage = $element->getImageRelationId() + ($element->isInSection() ? 6 : 0);
        $shapeId = md5($rIdObject . '_' . $rIdImage);
        $objectId = $element->getRelationId() + 1325353440;

        $style = $element->getStyle();
        $styleWriter = new ImageStyleWriter($xmlWriter, $style);

        if (!$this->withoutP) {
            $xmlWriter->startElement('w:p');
            $styleWriter->writeAlignment();
        }
        $this->writeCommentRangeStart();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:object');
        $xmlWriter->writeAttribute('w:dxaOrig', '15382');
        $xmlWriter->writeAttribute('w:dyaOrig', '5771');
        $xmlWriter->writeAttribute('w14:anchorId', '354C751A');

        $xmlWriter->startElement('v:shapetype');
        $xmlWriter->writeAttribute('id', '_x0000_t75');
        $xmlWriter->writeAttribute('coordsize', '21600,21600');
        $xmlWriter->writeAttribute('o:spt', $shapeId);
        $xmlWriter->writeAttribute('o:preferrelative', 't');
        $xmlWriter->writeAttribute('path', 'm@4@5l@4@11@9@11@9@5xe');
        $xmlWriter->writeAttribute('filled', 'f');
        $xmlWriter->writeAttribute('stroked', 'f');
        $xmlWriter->startElement('v:stroke');
        $xmlWriter->writeAttribute('joinstyle', 'miter');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:formulas');


        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'if lineDrawn pixelLineWidth 0');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'sum @0 1 0');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'sum 0 0 @1');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'prod @2 1 2');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'prod @3 21600 pixelWidth');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'prod @3 21600 pixelHeight');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'sum @0 0 1');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'prod @6 1 2');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'prod @7 21600 pixelWidth');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'sum @8 21600 0');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'prod @7 21600 pixelHeight');
        $xmlWriter->endElement();
        $xmlWriter->startElement('v:f');
        $xmlWriter->writeAttribute('eqn', 'sum @10 21600 0');
        $xmlWriter->endElement();
        $xmlWriter->fullEndElement();
        $xmlWriter->startElement('v:path');
        $xmlWriter->writeAttribute('o:extrusionok', 'f');
        $xmlWriter->writeAttribute('gradientshapeok', 't');
        $xmlWriter->writeAttribute('o:connecttype', 'rect');
        $xmlWriter->endElement();

        $xmlWriter->startElement('o:lock');
        $xmlWriter->writeAttribute('v:ext', 'edit');
        $xmlWriter->writeAttribute('aspectratio', 't');
        $xmlWriter->endElement();
        $xmlWriter->fullEndElement();


        // Icon
        $xmlWriter->startElement('v:shape');
        $xmlWriter->writeAttribute('id', $shapeId);
        $xmlWriter->writeAttribute('type', '#_x0000_t75');
        $xmlWriter->writeAttribute('style', 'width:368.4pt;height:129.5pt;mso-wrap-distance-right:0pt');
        $xmlWriter->writeAttribute('o:ole', '');

        $xmlWriter->startElement('v:imagedata');
        $xmlWriter->writeAttribute('r:id', 'rId' . $rIdImage);
        $xmlWriter->writeAttribute('o:title', '');
        $xmlWriter->endElement(); // v:imagedata

        $xmlWriter->endElement(); // v:shape

        // Object
        $xmlWriter->startElement('o:OLEObject');
        $xmlWriter->writeAttribute('Type', 'Embed');
        $xmlWriter->writeAttribute('ProgID', 'Excel.Sheet.12');
        $xmlWriter->writeAttribute('ShapeID', $shapeId);
        $xmlWriter->writeAttribute('DrawAspect', 'Content');
        $xmlWriter->writeAttribute('ObjectID', '_' . $objectId);
        $xmlWriter->writeAttribute('r:id', 'rId' . $rIdObject);
        $xmlWriter->endElement(); // o:OLEObject

        $xmlWriter->endElement(); // w:object
        $xmlWriter->endElement(); // w:r

        $this->endElementP(); // w:p
    }
}
