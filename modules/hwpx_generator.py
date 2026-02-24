"""
HWPX (OWPML) 전자책 생성기
KS X 6101 / OWPML 표준 - Hancom HWP Viewer / 한컴오피스 호환

레퍼런스: https://github.com/Canine89/hwpxskill (base 템플릿 구조 기반)

HWPUNIT = 1/100 pt  (100 units = 1pt)
  A4: 595.28pt × 841.89pt → 59528 × 84189 HWPUNIT
  1mm = 72/25.4 pt → 1mm = 283.46 HWPUNIT

ZIP 구조:
  mimetype                     (STORED, first)
  META-INF/container.xml
  version.xml
  settings.xml
  Contents/content.hpf
  Contents/header.xml
  Contents/section0.xml
  Preview/PrvText.txt

Namespaces:
  hs  = http://www.hancom.co.kr/hwpml/2011/section   (hs:sec 루트)
  hp  = http://www.hancom.co.kr/hwpml/2011/paragraph  (hp:p, hp:run, hp:t)
  hh  = http://www.hancom.co.kr/hwpml/2011/head       (header.xml 루트)
  hc  = http://www.hancom.co.kr/hwpml/2011/core       (hc:left 등)
  hv  = http://www.hancom.co.kr/hwpml/2011/version    (version.xml)
  ha  = http://www.hancom.co.kr/hwpml/2011/app        (settings.xml)
"""

import os
import re
import zipfile
import io
import xml.sax.saxutils as saxutils
from datetime import datetime, timezone
from config import load_config


def _esc(text):
    """XML 특수문자 이스케이프"""
    return saxutils.escape(str(text))


# ─── 단위 변환 ────────────────────────────────────────────────────
# HWPUNIT = 1/100 pt  →  100 HWPUNIT = 1pt
def _pt_to_hwp(pt):
    """pt → HWPUNIT"""
    return int(pt * 100)

def _mm_to_hwp(mm):
    """mm → HWPUNIT  (1mm = 72/25.4 pt)"""
    return int(mm * 72 / 25.4 * 100)


class EbookHwpxGenerator:
    """
    OWPML 기반 HWPX 전자책 생성기.
    Canine89/hwpxskill 의 base 템플릿 구조를 따릅니다.
    """

    def __init__(self, config=None):
        self.config = config or load_config()
        self.output_dir = self.config.get('output_dir', './static/output')
        os.makedirs(self.output_dir, exist_ok=True)
        self._para_id = 0

    def _new_pid(self):
        pid = self._para_id
        self._para_id += 1
        return str(pid)

    # ================================================================
    # mimetype
    # ================================================================
    def _mimetype(self):
        return b'application/hwp+zip'

    # ================================================================
    # META-INF/container.xml
    # ================================================================
    def _container_xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8'?>\n"
            '<ocf:container'
            ' xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container"'
            ' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">\n'
            '  <ocf:rootfiles>\n'
            '    <ocf:rootfile full-path="Contents/content.hpf"'
            ' media-type="application/hwpml-package+xml"/>\n'
            '    <ocf:rootfile full-path="Preview/PrvText.txt"'
            ' media-type="text/plain"/>\n'
            '  </ocf:rootfiles>\n'
            '</ocf:container>'
        )

    # ================================================================
    # version.xml  – 레퍼런스 템플릿과 동일한 포맷
    # ================================================================
    def _version_xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8'?>\n"
            '<hv:HCFVersion'
            ' xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version"'
            ' tagetApplication="WORDPROCESSOR"'
            ' major="5" minor="1" micro="1" buildNumber="0"'
            ' os="1" xmlVersion="1.5"'
            ' application="Hancom Office Hangul"'
            ' appVersion="13, 0, 0, 1408"/>'
        )

    # ================================================================
    # settings.xml
    # ================================================================
    def _settings_xml(self):
        return (
            "<?xml version='1.0' encoding='UTF-8'?>\n"
            '<ha:HWPApplicationSetting'
            ' xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"'
            ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">\n'
            '  <ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>\n'
            '</ha:HWPApplicationSetting>'
        )

    # ================================================================
    # Contents/content.hpf  (OPF 패키지 매니페스트)
    # ================================================================
    def _content_hpf(self, title):
        now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        return (
            "<?xml version='1.0' encoding='UTF-8'?>\n"
            '<opf:package'
            ' xmlns:opf="http://www.idpf.org/2007/opf/"'
            ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
            ' version="" unique-identifier="" id="">\n'
            '  <opf:metadata>\n'
            f'    <dc:title>{_esc(title)}</dc:title>\n'
            '    <dc:language>ko</dc:language>\n'
            f'    <opf:meta name="CreatedDate" content="text">{now}</opf:meta>\n'
            f'    <opf:meta name="ModifiedDate" content="text">{now}</opf:meta>\n'
            '    <opf:meta name="creator" content="text"/>\n'
            '    <opf:meta name="description" content="text"/>\n'
            '  </opf:metadata>\n'
            '  <opf:manifest>\n'
            '    <opf:item id="header" href="Contents/header.xml"'
            ' media-type="application/xml"/>\n'
            '    <opf:item id="section0" href="Contents/section0.xml"'
            ' media-type="application/xml"/>\n'
            '    <opf:item id="settings" href="settings.xml"'
            ' media-type="application/xml"/>\n'
            '  </opf:manifest>\n'
            '  <opf:spine>\n'
            '    <opf:itemref idref="header"/>\n'
            '    <opf:itemref idref="section0"/>\n'
            '  </opf:spine>\n'
            '</opf:package>'
        )

    # ================================================================
    # Contents/header.xml
    # 레퍼런스: version="1.5", itemCnt 필수
    # ================================================================
    def _header_xml(self, fs_hwp, hs_hwp, ss_hwp, small_hwp, ls_pct):
        return f"""<?xml version='1.0' encoding='UTF-8'?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"
         version="1.5" secCnt="1">

  <hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>

  <hh:refList>

    <!-- ── 폰트 ── -->
    <hh:fontfaces>
      <hh:fontface lang="HANGUL" itemCnt="3">
        <hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="2" weight="8" proportion="4" contrast="0"
                       strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>
        </hh:font>
        <hh:font id="1" face="함초롬바탕" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="2" weight="8" proportion="4" contrast="0"
                       strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>
        </hh:font>
        <hh:font id="2" face="맑은 고딕" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="2" weight="8" proportion="4" contrast="0"
                       strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>
        </hh:font>
      </hh:fontface>
      <hh:fontface lang="LATIN" itemCnt="2">
        <hh:font id="0" face="Arial" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="4" weight="5" proportion="4" contrast="0"
                       strokeVariation="2" armStyle="1" letterform="1" midline="2" xHeight="4"/>
        </hh:font>
        <hh:font id="1" face="Times New Roman" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="2" weight="5" proportion="6" contrast="5"
                       strokeVariation="2" armStyle="2" letterform="1" midline="2" xHeight="4"/>
        </hh:font>
      </hh:fontface>
      <hh:fontface lang="HANJA" itemCnt="2">
        <hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="2" weight="8" proportion="4" contrast="0"
                       strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>
        </hh:font>
        <hh:font id="1" face="함초롬바탕" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="2" weight="8" proportion="4" contrast="0"
                       strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>
        </hh:font>
      </hh:fontface>
      <hh:fontface lang="JAPANESE" itemCnt="2">
        <hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="2" weight="8" proportion="4" contrast="0"
                       strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>
        </hh:font>
        <hh:font id="1" face="함초롬바탕" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="2" weight="8" proportion="4" contrast="0"
                       strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>
        </hh:font>
      </hh:fontface>
      <hh:fontface lang="OTHER" itemCnt="1">
        <hh:font id="0" face="Arial" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="4" weight="5" proportion="4" contrast="0"
                       strokeVariation="2" armStyle="1" letterform="1" midline="2" xHeight="4"/>
        </hh:font>
      </hh:fontface>
      <hh:fontface lang="SYMBOL" itemCnt="1">
        <hh:font id="0" face="Symbol" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="4" weight="5" proportion="4" contrast="0"
                       strokeVariation="2" armStyle="1" letterform="1" midline="2" xHeight="4"/>
        </hh:font>
      </hh:fontface>
      <hh:fontface lang="USER" itemCnt="1">
        <hh:font id="0" face="Arial" type="TTF" isEmbedded="false">
          <hh:typeInfo familyType="4" weight="5" proportion="4" contrast="0"
                       strokeVariation="2" armStyle="1" letterform="1" midline="2" xHeight="4"/>
        </hh:font>
      </hh:fontface>
    </hh:fontfaces>

    <!-- ── 테두리/채우기 (itemCnt="2") ── -->
    <hh:borderFills itemCnt="2">
      <hh:borderFill id="0" themeType="NONE">
        <hc:fillBrush><hc:noFill/></hc:fillBrush>
        <hh:border>
          <hh:slash type="NONE"/>
          <hh:backSlash type="NONE"/>
          <hh:left type="NONE" width="0.1mm" color="#000000"/>
          <hh:right type="NONE" width="0.1mm" color="#000000"/>
          <hh:top type="NONE" width="0.1mm" color="#000000"/>
          <hh:bottom type="NONE" width="0.1mm" color="#000000"/>
        </hh:border>
      </hh:borderFill>
      <hh:borderFill id="1" themeType="NONE">
        <hc:fillBrush><hc:noFill/></hc:fillBrush>
        <hh:border>
          <hh:slash type="NONE"/>
          <hh:backSlash type="NONE"/>
          <hh:left type="NONE" width="0.1mm" color="#000000"/>
          <hh:right type="NONE" width="0.1mm" color="#000000"/>
          <hh:top type="NONE" width="0.1mm" color="#000000"/>
          <hh:bottom type="NONE" width="0.1mm" color="#000000"/>
        </hh:border>
      </hh:borderFill>
    </hh:borderFills>

    <!-- ── 글자 모양 (itemCnt="6") ──
         id=0: 본문  id=1: H1  id=2: H2  id=3: 라벨  id=4: bold  id=5: 목차 -->
    <hh:charProperties itemCnt="6">
      <hh:charPr id="0" height="{fs_hwp}" textColor="#000000" shadeColor="#FFFFFF"
                 useFontSpace="false" useKerning="false" symMark="NONE" borderFillIDRef="0">
        <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
      </hh:charPr>
      <hh:charPr id="1" height="{hs_hwp}" textColor="#1A1A2E" shadeColor="#FFFFFF"
                 useFontSpace="false" useKerning="false" symMark="NONE" borderFillIDRef="0">
        <hh:fontRef hangul="2" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:bold/>
      </hh:charPr>
      <hh:charPr id="2" height="{ss_hwp}" textColor="#1A1A2E" shadeColor="#FFFFFF"
                 useFontSpace="false" useKerning="false" symMark="NONE" borderFillIDRef="0">
        <hh:fontRef hangul="2" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:bold/>
      </hh:charPr>
      <hh:charPr id="3" height="{small_hwp}" textColor="#646464" shadeColor="#FFFFFF"
                 useFontSpace="false" useKerning="false" symMark="NONE" borderFillIDRef="0">
        <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
      </hh:charPr>
      <hh:charPr id="4" height="{fs_hwp}" textColor="#000000" shadeColor="#FFFFFF"
                 useFontSpace="false" useKerning="false" symMark="NONE" borderFillIDRef="0">
        <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:bold/>
      </hh:charPr>
      <hh:charPr id="5" height="{fs_hwp}" textColor="#000000" shadeColor="#FFFFFF"
                 useFontSpace="false" useKerning="false" symMark="NONE" borderFillIDRef="0">
        <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
      </hh:charPr>
    </hh:charProperties>

    <!-- ── 탭 속성 (itemCnt="2") ── -->
    <hh:tabProperties itemCnt="2">
      <hh:tabPr id="0" autoTabLeft="false" autoTabRight="false"/>
      <hh:tabPr id="1" autoTabLeft="false" autoTabRight="true"/>
    </hh:tabProperties>

    <!-- ── 번호 매기기 (itemCnt="1") ── -->
    <hh:numberings itemCnt="1">
      <hh:numbering id="0" start="1">
        <hh:paraHead level="1" start="1" unified="false">
          <hh:numFormat/>
        </hh:paraHead>
      </hh:numbering>
    </hh:numberings>

    <!-- ── 문단 모양 (itemCnt="5") ──
         id=0: 본문  id=1: H1 단락  id=2: H2 단락  id=3: 목차  id=4: 들여쓰기 -->
    <hh:paraProperties itemCnt="5">
      <hh:paraPr id="0" tabPrIDRef="0" condense="0" fontLineHeight="false"
                 snapToGrid="true" suppressLineNumbers="false" checked="false">
        <hh:align horizontal="LEFT" vertical="BASELINE"/>
        <hh:lineSpacing type="PERCENT" value="{ls_pct}"/>
        <hh:margin>
          <hc:left value="0" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="0" unit="HWPUNIT"/>
          <hc:next value="200" unit="HWPUNIT"/>
          <hc:indent value="0" unit="HWPUNIT"/>
        </hh:margin>
        <hh:border borderFillIDRef="0" offsetLeft="0" offsetRight="0"
                   offsetTop="0" offsetBottom="0"/>
      </hh:paraPr>
      <hh:paraPr id="1" tabPrIDRef="1" condense="0" fontLineHeight="false"
                 snapToGrid="true" suppressLineNumbers="false" checked="false">
        <hh:align horizontal="LEFT" vertical="BASELINE"/>
        <hh:lineSpacing type="PERCENT" value="{ls_pct}"/>
        <hh:margin>
          <hc:left value="0" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="1000" unit="HWPUNIT"/>
          <hc:next value="600" unit="HWPUNIT"/>
          <hc:indent value="0" unit="HWPUNIT"/>
        </hh:margin>
        <hh:border borderFillIDRef="0" offsetLeft="0" offsetRight="0"
                   offsetTop="0" offsetBottom="0"/>
      </hh:paraPr>
      <hh:paraPr id="2" tabPrIDRef="0" condense="0" fontLineHeight="false"
                 snapToGrid="true" suppressLineNumbers="false" checked="false">
        <hh:align horizontal="LEFT" vertical="BASELINE"/>
        <hh:lineSpacing type="PERCENT" value="{ls_pct}"/>
        <hh:margin>
          <hc:left value="0" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="600" unit="HWPUNIT"/>
          <hc:next value="300" unit="HWPUNIT"/>
          <hc:indent value="0" unit="HWPUNIT"/>
        </hh:margin>
        <hh:border borderFillIDRef="0" offsetLeft="0" offsetRight="0"
                   offsetTop="0" offsetBottom="0"/>
      </hh:paraPr>
      <hh:paraPr id="3" tabPrIDRef="0" condense="0" fontLineHeight="false"
                 snapToGrid="true" suppressLineNumbers="false" checked="false">
        <hh:align horizontal="LEFT" vertical="BASELINE"/>
        <hh:lineSpacing type="PERCENT" value="{ls_pct}"/>
        <hh:margin>
          <hc:left value="0" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="0" unit="HWPUNIT"/>
          <hc:next value="100" unit="HWPUNIT"/>
          <hc:indent value="0" unit="HWPUNIT"/>
        </hh:margin>
        <hh:border borderFillIDRef="0" offsetLeft="0" offsetRight="0"
                   offsetTop="0" offsetBottom="0"/>
      </hh:paraPr>
      <hh:paraPr id="4" tabPrIDRef="0" condense="0" fontLineHeight="false"
                 snapToGrid="true" suppressLineNumbers="false" checked="false">
        <hh:align horizontal="LEFT" vertical="BASELINE"/>
        <hh:lineSpacing type="PERCENT" value="{ls_pct}"/>
        <hh:margin>
          <hc:left value="400" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="0" unit="HWPUNIT"/>
          <hc:next value="100" unit="HWPUNIT"/>
          <hc:indent value="-400" unit="HWPUNIT"/>
        </hh:margin>
        <hh:border borderFillIDRef="0" offsetLeft="0" offsetRight="0"
                   offsetTop="0" offsetBottom="0"/>
      </hh:paraPr>
    </hh:paraProperties>

    <!-- ── 스타일 (itemCnt="6") ── -->
    <hh:styles itemCnt="6">
      <hh:style id="0" type="PARA" name="바탕글" engName="Normal"
                paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0"
                langID="1042" lockForm="false"/>
      <hh:style id="1" type="PARA" name="본문" engName="Body"
                paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="1"
                langID="1042" lockForm="false"/>
      <hh:style id="2" type="PARA" name="개요 1" engName="Outline 1"
                paraPrIDRef="1" charPrIDRef="1" nextStyleIDRef="0"
                langID="1042" lockForm="false"/>
      <hh:style id="3" type="PARA" name="개요 2" engName="Outline 2"
                paraPrIDRef="2" charPrIDRef="2" nextStyleIDRef="1"
                langID="1042" lockForm="false"/>
      <hh:style id="4" type="PARA" name="목차" engName="TOC"
                paraPrIDRef="3" charPrIDRef="5" nextStyleIDRef="4"
                langID="1042" lockForm="false"/>
      <hh:style id="5" type="PARA" name="목록 들여쓰기" engName="List Indent"
                paraPrIDRef="4" charPrIDRef="0" nextStyleIDRef="5"
                langID="1042" lockForm="false"/>
    </hh:styles>

  </hh:refList>

</hh:head>"""

    # ================================================================
    # Contents/section0.xml
    # 레퍼런스: hs:sec > hp:p(secPr+colPr) > hp:p(content)...
    # ================================================================

    def _secpr(self, pw, ph, ml, mr, mt, mb):
        """
        섹션 속성 XML – 레퍼런스 base 템플릿 속성 구조와 동일
        pw/ph = 페이지 너비/높이 (HWPUNIT)
        ml/mr/mt/mb = 여백 (HWPUNIT)
        """
        hdr = _mm_to_hwp(15)
        ftr = _mm_to_hwp(15)
        return (
            f'<hp:secPr id="0" textDirection="HORIZONTAL" spaceColumns="1134"'
            f' tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT"'
            f' outlineShapeIDRef="0" memoShapeIDRef="0"'
            f' textVerticalWidthHead="0" masterPageCnt="0">\n'
            f'      <hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>\n'
            f'      <hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>\n'
            f'      <hp:visibility hideFirstHeader="0" hideFirstFooter="0"'
            f' hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL"'
            f' hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>\n'
            f'      <hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>\n'
            f'      <hp:pagePr landscape="WIDELY" width="{pw}" height="{ph}" gutterType="LEFT_ONLY">\n'
            f'        <hp:margin header="{hdr}" footer="{ftr}" gutter="0"'
            f' left="{ml}" right="{mr}" top="{mt}" bottom="{mb}"/>\n'
            f'      </hp:pagePr>\n'
            f'      <hp:footNotePr>\n'
            f'        <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>\n'
            f'        <hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>\n'
            f'        <hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>\n'
            f'        <hp:numbering type="CONTINUOUS" newNum="1"/>\n'
            f'        <hp:placement place="EACH_COLUMN" beneathText="0"/>\n'
            f'      </hp:footNotePr>\n'
            f'      <hp:endNotePr>\n'
            f'        <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>\n'
            f'        <hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>\n'
            f'        <hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>\n'
            f'        <hp:numbering type="CONTINUOUS" newNum="1"/>\n'
            f'        <hp:placement place="END_OF_DOCUMENT" beneathText="0"/>\n'
            f'      </hp:endNotePr>\n'
            f'      <hp:pageBorderFill type="BOTH" borderFillIDRef="0"'
            f' textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n'
            f'        <hp:offset left="1417" right="1417" top="1417" bottom="1417"/>\n'
            f'      </hp:pageBorderFill>\n'
            f'      <hp:pageBorderFill type="EVEN" borderFillIDRef="0"'
            f' textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n'
            f'        <hp:offset left="1417" right="1417" top="1417" bottom="1417"/>\n'
            f'      </hp:pageBorderFill>\n'
            f'      <hp:pageBorderFill type="ODD" borderFillIDRef="0"'
            f' textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n'
            f'        <hp:offset left="1417" right="1417" top="1417" bottom="1417"/>\n'
            f'      </hp:pageBorderFill>\n'
            f'    </hp:secPr>'
        )

    def _setup_para(self, secpr_xml):
        """
        섹션의 첫 번째 단락 – secPr + colPr 포함 (레퍼런스 구조와 동일)
        run1: secPr + ctrl(colPr)
        run2: 빈 텍스트
        """
        pid = self._new_pid()
        return (
            f'  <hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">\n'
            f'    <hp:run charPrIDRef="0">\n'
            f'      {secpr_xml}\n'
            f'      <hp:ctrl>\n'
            f'        <hp:colPr id="0" type="NEWSPAPER" layout="LEFT"'
            f' colCount="1" sameSz="1" sameGap="0"/>\n'
            f'      </hp:ctrl>\n'
            f'    </hp:run>\n'
            f'    <hp:run charPrIDRef="0">\n'
            f'      <hp:t/>\n'
            f'    </hp:run>\n'
            f'    <hp:linesegarray>\n'
            f'      <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
            f' baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>\n'
            f'    </hp:linesegarray>\n'
            f'  </hp:p>\n'
        )

    def _make_para(self, text, para_pr_id=0, char_pr_id=0,
                   page_break=False, style_id=0):
        """일반 콘텐츠 단락 hp:p 생성"""
        pid = self._new_pid()
        pb = '1' if page_break else '0'

        if not text:
            return (
                f'  <hp:p id="{pid}" paraPrIDRef="{para_pr_id}"'
                f' styleIDRef="{style_id}" pageBreak="{pb}"'
                f' columnBreak="0" merged="0">\n'
                f'    <hp:run charPrIDRef="{char_pr_id}"/>\n'
                f'  </hp:p>\n'
            )
        return (
            f'  <hp:p id="{pid}" paraPrIDRef="{para_pr_id}"'
            f' styleIDRef="{style_id}" pageBreak="{pb}"'
            f' columnBreak="0" merged="0">\n'
            f'    <hp:run charPrIDRef="{char_pr_id}">\n'
            f'      <hp:t>{_esc(text)}</hp:t>\n'
            f'    </hp:run>\n'
            f'  </hp:p>\n'
        )

    # ── 편의 단락 메서드 ─────────────────────────────────────────
    def _h1(self, text, page_break=False):
        return self._make_para(text, para_pr_id=1, char_pr_id=1,
                               style_id=2, page_break=page_break)

    def _h2(self, text):
        return self._make_para(text, para_pr_id=2, char_pr_id=2, style_id=3)

    def _body(self, text):
        return self._make_para(text, para_pr_id=0, char_pr_id=0, style_id=1)

    def _label(self, text):
        return self._make_para(text, para_pr_id=0, char_pr_id=3, style_id=0)

    def _bold(self, text):
        return self._make_para(text, para_pr_id=0, char_pr_id=4, style_id=1)

    def _toc(self, text):
        return self._make_para(text, para_pr_id=3, char_pr_id=5, style_id=4)

    def _bullet(self, text):
        return self._make_para('• ' + text, para_pr_id=4, char_pr_id=0,
                               style_id=5)

    def _empty(self):
        return self._make_para('', para_pr_id=0, char_pr_id=0, style_id=0)

    # ================================================================
    # section0.xml 전체 조립
    # ================================================================
    def _section_xml(self, content_paras_xml, pw, ph, ml, mr, mt, mb):
        secpr_xml = self._secpr(pw, ph, ml, mr, mt, mb)
        setup = self._setup_para(secpr_xml)
        return (
            "<?xml version='1.0' encoding='UTF-8'?>\n"
            '<hs:sec'
            ' xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"'
            ' xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"'
            ' xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph"'
            ' xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"'
            ' xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"'
            ' xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"'
            ' xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history"'
            ' xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"'
            ' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"'
            ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
            ' xmlns:opf="http://www.idpf.org/2007/opf/"'
            ' xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"'
            ' xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"'
            ' xmlns:epub="http://www.idpf.org/2007/ops"'
            ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">\n'
            + setup
            + content_paras_xml
            + '</hs:sec>'
        )

    # ================================================================
    # 메인 generate 메서드
    # ================================================================
    def generate(self, ebook_data):
        self._para_id = 0

        cfg = self.config
        fs  = cfg.get('pdf_font_size', 11)
        hs  = cfg.get('pdf_heading_size', 16)
        ss  = cfg.get('pdf_subheading_size', 13)
        ls  = cfg.get('pdf_line_spacing', 1.6)
        ml  = cfg.get('pdf_margin_left', 60)
        mr  = cfg.get('pdf_margin_right', 60)
        mt  = cfg.get('pdf_margin_top', 72)
        mb  = cfg.get('pdf_margin_bottom', 72)

        # HWPUNIT 변환 (100 HWPUNIT = 1pt)
        fs_hwp    = int(fs * 100)
        hs_hwp    = int(hs * 100)
        ss_hwp    = int(ss * 100)
        small_hwp = int(max(fs - 1.5, 8) * 100)
        ls_pct    = int(ls * 100)

        # 여백: pt → HWPUNIT
        ml_hwp = _pt_to_hwp(ml)
        mr_hwp = _pt_to_hwp(mr)
        mt_hwp = _pt_to_hwp(mt)
        mb_hwp = _pt_to_hwp(mb)

        # A4 portrait (595.28pt × 841.86pt → HWPUNIT)
        page_w = 59528
        page_h = 84186

        book_info = ebook_data.get('book_info', {})
        title     = book_info.get('book_title', '전자책')
        subtitle  = book_info.get('subtitle', '')
        safe_title = re.sub(r'[^\w가-힣\s\-]', '', title)[:50].strip()
        filename   = f"{safe_title}.hwpx"
        filepath   = os.path.join(self.output_dir, filename)

        # ── 콘텐츠 단락 구성 ──────────────────────────────────────
        paras = []

        # 표지
        paras.append(self._h1(title))
        if subtitle:
            paras.append(self._h2(subtitle))
        paras.append(self._empty())
        analysis = ebook_data.get('analysis', {})
        if analysis.get('target_reader'):
            paras.append(self._body(f"대상 독자: {analysis['target_reader']}"))

        # 목차
        paras.append(self._h1('목  차', page_break=True))
        chapters = book_info.get('chapters', [])
        for ch in chapters:
            phase    = ch.get('phase', '')
            num      = ch.get('chapter_num', '')
            ch_title = ch.get('title', '')
            paras.append(self._toc(f"[{phase}]  CHAPTER {num}  {ch_title}"))
        paras.append(self._empty())

        # 가치 요약
        if analysis:
            paras.append(self._h1('이 책이 주는 가치', page_break=True))
            problem = analysis.get('problem_solved', {})
            for lbl, key in [('절약 시간', 'time'),
                              ('비용 절감', 'money'),
                              ('감정적 해방', 'emotion')]:
                val = problem.get(key, '')
                if val:
                    paras.append(self._h2(lbl))
                    paras.append(self._body(val))
            if analysis.get('why_pay'):
                paras.append(self._h2('왜 이 책에 투자해야 하는가'))
                paras.append(self._body(analysis['why_pay']))

        # 챕터 본문
        for i, ch_data in enumerate(ebook_data.get('chapters_content', [])):
            chapter = ch_data.get('chapter', {})
            content = ch_data.get('content', '')

            paras.append(self._label(
                f"CHAPTER {i+1}  ·  {chapter.get('phase', '')}"
            ))
            paras.append(self._h1(chapter.get('title', ''), page_break=True))

            before = chapter.get('before_state', '')
            after  = chapter.get('after_state', '')
            if before:
                paras.append(self._body(f"읽기 전: {before}"))
            if after:
                paras.append(self._bold(f"읽고 난 후: {after}"))
            paras.append(self._empty())

            for line in content.split('\n'):
                stripped = line.strip()
                if not stripped:
                    paras.append(self._empty())
                    continue
                m = re.match(r'^={2,}\s*(.+?)\s*={2,}$', stripped)
                if m:
                    paras.append(self._h2(m.group(1)))
                    continue
                bm = re.match(r'^\[([^\]]{2,20})\]\s*(.*)', stripped)
                if bm:
                    label_txt = bm.group(1)
                    rest = bm.group(2).strip()
                    paras.append(self._bold(f'[ {label_txt} ]'))
                    if rest:
                        paras.append(self._body(rest))
                    continue
                if re.match(r'^[-•●▶►✓]\s+', stripped):
                    paras.append(self._bullet(stripped[2:].strip()))
                    continue
                m2 = re.match(r'^(\d+)[.)]\s+(.+)', stripped)
                if m2:
                    paras.append(self._bullet(f"{m2.group(1)}. {m2.group(2)}"))
                    continue
                paras.append(self._body(stripped))

        # 마케팅 부록
        marketing = ebook_data.get('marketing', {})
        if marketing:
            paras.append(self._h1('부록: 이 책에 대하여', page_break=True))
            if marketing.get('sales_copy'):
                paras.append(self._h2('판매 소개문'))
                paras.append(self._body(marketing['sales_copy']))
            value = marketing.get('value_summary', {})
            if value:
                paras.append(self._h2('독자에게 주는 가치'))
                for key, lbl in [('time_saved', '절약 시간'),
                                  ('money_saved', '비용 절감'),
                                  ('mistakes_prevented', '방지 실수')]:
                    if value.get(key):
                        paras.append(self._bullet(f"{lbl}: {value[key]}"))

        # ── XML 조립 ──────────────────────────────────────────────
        content_xml = ''.join(paras)
        section_xml = self._section_xml(
            content_xml, page_w, page_h, ml_hwp, mr_hwp, mt_hwp, mb_hwp
        )
        header_xml = self._header_xml(fs_hwp, hs_hwp, ss_hwp, small_hwp, ls_pct)
        preview    = re.sub(r'<[^>]+>', '', content_xml)[:500]

        # ── HWPX ZIP 패키징 ──────────────────────────────────────
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            # mimetype: STORED, 첫 번째 항목
            mime_info = zipfile.ZipInfo('mimetype')
            mime_info.compress_type = zipfile.ZIP_STORED
            zf.writestr(mime_info, self._mimetype())

            zf.writestr('META-INF/container.xml',
                        self._container_xml().encode('utf-8'))
            zf.writestr('version.xml',
                        self._version_xml().encode('utf-8'))
            zf.writestr('settings.xml',
                        self._settings_xml().encode('utf-8'))
            zf.writestr('Contents/content.hpf',
                        self._content_hpf(title).encode('utf-8'))
            zf.writestr('Contents/header.xml',
                        header_xml.encode('utf-8'))
            zf.writestr('Contents/section0.xml',
                        section_xml.encode('utf-8'))
            zf.writestr('Preview/PrvText.txt',
                        preview.encode('utf-8'))

        with open(filepath, 'wb') as f:
            f.write(buf.getvalue())

        return filepath, filename
