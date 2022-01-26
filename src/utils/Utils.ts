import { AppConfig } from "./AppConfig";

/** Logs debug traces to the console, if enabled. */
export function trace(text: string) {
  if (AppConfig.trace) {
    console.log(text);
  }
}

/** Logs error messages to the console. */
export function error(text: string) {
  console.error(text);
}

/** Click handler for button to insert a foot- or and endnote in the Word doc */
export async function onInsertNoteClicked(data: any, which: string) {
  return Word.run(async function (context) {
    const insertEndnote = which === "footnote" ? false : true;
    const range = context.document.getSelection();
    await context.sync();

    insertNote(context, range, insertEndnote /* useEndnote */, data);

    await context.sync();
  });
}

/**
 * Inserts either a footnote or an endnote
 * @param context : the current Word context
 * @param selection : the current text selection (Range) in the Word doc
 * @param useEndnote : true to insert an endnote, false to insert a footnote
 */
function insertNote(context: Word.RequestContext, selectionToInsertAfter: Word.Range, useEndnote: boolean, data: any) {
  const noteText = createNoteText(data.description, data.reference);
  const xml = useEndnote ? createEndnoteXml(noteText) : createFootnoteXml(noteText);
  selectionToInsertAfter.insertOoxml(xml, "After");
}

/** Create the OOXML text for the footnote/endnote text. */
function createNoteText(description: string, reference: string): string {
  description = escapeHtml(description + "\n");
  reference = escapeHtml(reference);

  let xml = `<w:t xml:space="preserve">${description} [${reference}]</w:t>`;
  return xml;
}

/** Escape non alpha-numeric chars for safe inclusion in HTML */
function escapeHtml(text: string) {
  return text ? text.replace(/[^0-9A-Za-z ]/g, (char) => "&#" + char.charCodeAt(0) + ";") : "";
}

/** Creates the OOXML text needed to add a footnote to the Word document. */
function createFootnoteXml(noteText: string): string {
  // Note: the <?xml?> tag must be on the first line, or Word won't accept it (console error: Unhandled promise rejection)
  var xml: string = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<?mso-application progid="Word.Document"?>
		<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
			<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
				<pkg:xmlData>
					<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
						<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
					</Relationships>
				</pkg:xmlData>
			</pkg:part>
		
			<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
				<pkg:xmlData>
					<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14">
						<w:body>
							<w:p w:rsidR="00000000" w:rsidRDefault="00FD481E">
								<w:r w:rsidRPr="24EDB37C">
									<w:rPr>
										<w:rStyle w:val="FootnoteReference" />
										<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri" />
										<w:color w:val="333333" />
										<w:sz w:val="21" />
										<w:szCs w:val="21" />
									</w:rPr>
									<w:footnoteReference w:id="1" />
								</w:r>
							</w:p>
							<w:sectPr w:rsidR="00000000">
								<w:pgSz w:w="12240" w:h="15840" />
								<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0" />
								<w:cols w:space="720" />
							</w:sectPr>
						</w:body>
					</w:document>
				</pkg:xmlData>
			</pkg:part>
		
			<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="256">
				<pkg:xmlData>
					<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
						<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
						<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml" />
					</Relationships>
				</pkg:xmlData>
			</pkg:part>
		
			<pkg:part pkg:name="/word/footnotes.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml">
				<pkg:xmlData>
					<w:footnotes xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14">
						<w:footnote w:type="separator" w:id="-1">
							<w:p w:rsidR="00FD481E" w:rsidRDefault="00FD481E">
								<w:pPr>
									<w:spacing w:after="0" w:line="240" w:lineRule="auto" />
								</w:pPr>
								<w:r>
									<w:separator />
								</w:r>
							</w:p>
						</w:footnote>
						<w:footnote w:type="continuationSeparator" w:id="0">
							<w:p w:rsidR="00FD481E" w:rsidRDefault="00FD481E">
								<w:pPr>
									<w:spacing w:after="0" w:line="240" w:lineRule="auto" />
								</w:pPr>
								<w:r>
									<w:continuationSeparator />
								</w:r>
							</w:p>
						</w:footnote>
						<w:footnote w:id="1">
							<w:p w:rsidR="00FD481E" w:rsidRDefault="00FD481E" w:rsidP="24EDB37C">
								<w:pPr>
									<w:pStyle w:val="FootnoteText" />
								</w:pPr>
								<w:r w:rsidRPr="24EDB37C">
									<w:rPr>
										<w:rStyle w:val="FootnoteReference" />
									</w:rPr>
									<w:footnoteRef />
								</w:r>
								<w:r>
									${noteText}
								</w:r>
							</w:p>
						</w:footnote>
					</w:footnotes>
				</pkg:xmlData>
			</pkg:part>
		
			<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
				<pkg:xmlData>
					<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex">
						<w:style w:type="character" w:styleId="FootnoteReference">
							<w:name w:val="footnote reference" />
							<w:basedOn w:val="DefaultParagraphFont" />
							<w:uiPriority w:val="99" />
							<w:semiHidden />
							<w:unhideWhenUsed />
							<w:rPr>
								<w:vertAlign w:val="superscript" />
							</w:rPr>
						</w:style>
						<w:style w:type="character" w:customStyle="1" w:styleId="FootnoteTextChar">
							<w:name w:val="Footnote Text Char" />
							<w:basedOn w:val="DefaultParagraphFont" />
							<w:link w:val="FootnoteText" />
							<w:uiPriority w:val="99" />
							<w:semiHidden />
							<w:rPr>
								<w:sz w:val="20" />
								<w:szCs w:val="20" />
							</w:rPr>
						</w:style>
						<w:style w:type="paragraph" w:styleId="FootnoteText">
							<w:name w:val="footnote text" />
							<w:basedOn w:val="Normal" />
							<w:link w:val="FootnoteTextChar" />
							<w:uiPriority w:val="99" />
							<w:semiHidden />
							<w:unhideWhenUsed />
							<w:pPr>
								<w:spacing w:after="0" w:line="240" w:lineRule="auto" />
							</w:pPr>
							<w:rPr>
								<w:sz w:val="20" />
								<w:szCs w:val="20" />
							</w:rPr>
						</w:style>
					</w:styles>
				</pkg:xmlData>
			</pkg:part>
		
		</pkg:package>`;

  return xml;
}

/** Creates the OOXML text needed to add an endnote to the Word document. */
function createEndnoteXml(noteText: string): string {
  // Note: the <?xml?> tag must be on the first line, or Word won't accept it (console error: Unhandled promise rejection)
  var xml: string = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<?mso-application progid="Word.Document"?>
		<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
			<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
				<pkg:xmlData>
					<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
						<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
					</Relationships>
				</pkg:xmlData>
			</pkg:part>
			<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
				<pkg:xmlData>
					<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14">
						<w:body>
							<w:p w:rsidR="00000000" w:rsidRDefault="00FD481E">
								<w:r w:rsidRPr="24EDB37C">
									<w:rPr>
										<w:rStyle w:val="EndnoteReference" />
									</w:rPr>
									<w:endnoteReference w:id="1" />
								</w:r>
							</w:p>
							<w:sectPr w:rsidR="00000000">
								<w:pgSz w:w="12240" w:h="15840" />
								<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:ender="720" w:gutter="0" />
								<w:cols w:space="720" />
							</w:sectPr>
						</w:body>
					</w:document>
				</pkg:xmlData>
			</pkg:part>
			<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="256">
				<pkg:xmlData>
					<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
						<Relationship Id="rId1" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
						<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml" />
					</Relationships>
				</pkg:xmlData>
			</pkg:part>
			<pkg:part pkg:name="/word/endnotes.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml">
				<pkg:xmlData>
					<w:endnotes xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14">
						<w:endnote w:type="separator" w:id="-1">
							<w:p w:rsidR="00FD481E" w:rsidRDefault="00FD481E">
								<w:pPr>
									<w:spacing w:after="0" w:line="240" w:lineRule="auto" />
								</w:pPr>
								<w:r>
									<w:separator />
								</w:r>
							</w:p>
						</w:endnote>
						<w:endnote w:type="continuationSeparator" w:id="0">
							<w:p w:rsidR="00FD481E" w:rsidRDefault="00FD481E">
								<w:pPr>
									<w:spacing w:after="0" w:line="240" w:lineRule="auto" />
								</w:pPr>
								<w:r>
									<w:continuationSeparator />
								</w:r>
							</w:p>
						</w:endnote>
						<w:endnote w:id="1">
							<w:p>
								<w:pPr>
									<w:pStyle w:val="EndnoteText"/>
								</w:pPr>
								<w:r>
									<w:rPr>
										<w:rStyle w:val="EndnoteReference"/>
									</w:rPr>
									<w:endnoteRef/>
								</w:r>
								<w:r>
									${noteText}
								</w:r>
							</w:p>
						</w:endnote>
					</w:endnotes>
				</pkg:xmlData>
			</pkg:part>
			<pkg:part pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" pkg:name="/word/styles.xml">
				<pkg:xmlData>
					<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
						<w:style w:default="1" w:styleId="Normal" w:type="paragraph">
							<w:name w:val="Normal"/>
							<w:qFormat/>
						</w:style>
						<w:style w:default="1" w:styleId="DefaultParagraphFont" w:type="character">
							<w:name w:val="Default Paragraph Font"/>
							<w:uiPriority w:val="1"/>
							<w:semiHidden/>
							<w:unhideWhenUsed/>
						</w:style>
						<w:style w:default="1" w:styleId="NoList" w:type="numbering">
							<w:name w:val="No List"/>
							<w:uiPriority w:val="99"/>
							<w:semiHidden/>
							<w:unhideWhenUsed/>
						</w:style>
						<w:style w:styleId="EndnoteText" w:type="paragraph">
							<w:name w:val="endnote text"/>
							<w:basedOn w:val="Normal"/>
							<w:link w:val="EndnoteTextChar"/>
							<w:uiPriority w:val="99"/>
							<w:semiHidden/>
							<w:unhideWhenUsed/>
							<w:rPr>
								<w:sz w:val="20"/>
								<w:szCs w:val="20"/>
							</w:rPr>
						</w:style>
						<w:style w:customStyle="1" w:styleId="EndnoteTextChar" w:type="character">
							<w:name w:val="Endnote Text Char"/>
							<w:basedOn w:val="DefaultParagraphFont"/>
							<w:link w:val="EndnoteText"/>
							<w:uiPriority w:val="99"/>
							<w:semiHidden/>
							<w:rPr>
								<w:sz w:val="20"/>
								<w:szCs w:val="20"/>
							</w:rPr>
						</w:style>
						<w:style w:styleId="EndnoteReference" w:type="character">
							<w:name w:val="endnote reference"/>
							<w:basedOn w:val="DefaultParagraphFont"/>
							<w:uiPriority w:val="99"/>
							<w:semiHidden/>
							<w:unhideWhenUsed/>
							<w:rPr>
								<w:vertAlign w:val="superscript"/>
							</w:rPr>
						</w:style>
					</w:styles>
				</pkg:xmlData>
			</pkg:part>
		</pkg:package>`;

  return xml;
}
