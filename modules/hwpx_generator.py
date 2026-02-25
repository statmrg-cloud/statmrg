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

    <!-- ── 문단 모양 (itemCnt="6") ──
         id=0: 본문  id=1: H1 단락  id=2: H2 단락  id=3: 목차  id=4: 들여쓰기  id=5: 가운데(footer용) -->
    <hh:paraProperties itemCnt="6">
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
      <hh:paraPr id="5" tabPrIDRef="0" condense="0" fontLineHeight="false"
                 snapToGrid="true" suppressLineNumbers="false" checked="false">
        <hh:align horizontal="CENTER" vertical="BASELINE"/>
        <hh:lineSpacing type="PERCENT" value="{ls_pct}"/>
        <hh:margin>
          <hc:left value="0" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="0" unit="HWPUNIT"/>
          <hc:next value="0" unit="HWPUNIT"/>
          <hc:indent value="0" unit="HWPUNIT"/>
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
        섹션의 첫 번째 단락 – secPr + colPr + pageNum 포함
        run1: secPr + ctrl(colPr) + ctrl(pageNum 쪽번호매기기)
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
            f'      <hp:ctrl>\n'
            f'        <hp:pageNum pos="BOTTOM_CENTER" formatType="DIGIT" sideChar="-"/>\n'
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

    @staticmethod
    def _calc_content_height_pt(text, fs=11, hs_size=16, ss_size=13, ls=1.6,
                                 page_width_pt=475.28):
        """챕터 본문의 총 높이(pt)를 계산 — 렌더링 파싱 로직을 정확히 미러링.
        각 단락 스타일(H2, bold, body, bullet, empty)의 폰트 크기, 줄간격,
        단락 마진(prev/next)을 반영하여 실제 렌더링 높이를 시뮬레이션한다.
        """
        if not text:
            return 0.0

        # 줄당 전각 글자 수 (한글 기준: 1em = font_size pt)
        chars_per_line = int(page_width_pt / fs)
        chars_per_line_ss = int(page_width_pt / ss_size)

        def _wrap_lines(s, cpl=chars_per_line):
            return max(1, -(-len(s) // cpl))

        # 단락 높이 계산 함수 (줄 수 * 줄높이 + prev마진 + next마진)
        def _body_height(n_lines):
            """본문 단락 (paraPr 0): next=2pt"""
            return n_lines * (fs * ls) + 2.0

        def _h2_height(n_lines):
            """H2 단락 (paraPr 2): prev=6pt, next=3pt"""
            return 6.0 + n_lines * (ss_size * ls) + 3.0

        def _bold_height(n_lines):
            """볼드 단락 (paraPr 0): next=2pt"""
            return n_lines * (fs * ls) + 2.0

        def _bullet_height(n_lines):
            """불릿 단락 (paraPr 4): next=1pt"""
            return n_lines * (fs * ls) + 1.0

        def _empty_height():
            """빈 단락: 1줄 높이 + next=2pt"""
            return fs * ls + 2.0

        total_pt = 0.0
        for raw_line in text.split('\n'):
            stripped = raw_line.strip()
            if not stripped:
                total_pt += _empty_height()
                continue
            # == 소제목 == → H2
            if re.match(r'^={2,}\s*(.+?)\s*={2,}$', stripped):
                title_text = re.match(r'^={2,}\s*(.+?)\s*={2,}$', stripped).group(1)
                total_pt += _h2_height(_wrap_lines(title_text, chars_per_line_ss))
                continue
            # [라벨] 텍스트 → bold + body
            bm = re.match(r'^\[([^\]]{2,20})\]\s*(.*)', stripped)
            if bm:
                label_txt = f'[ {bm.group(1)} ]'
                total_pt += _bold_height(_wrap_lines(label_txt))
                rest = bm.group(2).strip()
                if rest:
                    total_pt += _body_height(_wrap_lines(rest))
                continue
            # 불릿
            if re.match(r'^[-•●▶►✓]\s+', stripped):
                total_pt += _bullet_height(_wrap_lines('• ' + stripped[2:].strip()))
                continue
            # 번호 리스트
            m2 = re.match(r'^(\d+)[.)]\s+(.+)', stripped)
            if m2:
                total_pt += _bullet_height(_wrap_lines(f'• {m2.group(1)}. {m2.group(2)}'))
                continue
            # 일반 본문
            total_pt += _body_height(_wrap_lines(stripped))
        return total_pt

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

        # 목차 (추정 페이지번호 포함)
        paras.append(self._h1('목  차', page_break=True))
        chapters = book_info.get('chapters', [])

        # ── 페이지번호 추정 (pt 기반 높이 시뮬레이션) ──
        # A4 본문 영역: 너비 595.28-60-60=475.28pt, 높이 841.86-72-72=697.86pt
        PAGE_H = 697.86  # 본문 영역 높이 (pt)
        PAGE_W = 475.28  # 본문 영역 너비 (pt)
        n_chapters = len(chapters)

        # H1 높이: 16pt × 1.6 + prev=10pt + next=6pt
        H1_HEIGHT = hs * ls + 10.0 + 6.0
        # H2 높이: 13pt × 1.6 + prev=6pt + next=3pt
        H2_HEIGHT = ss * ls + 6.0 + 3.0
        # 본문 줄 높이: 11pt × 1.6 + next=2pt
        BODY_LINE = fs * ls + 2.0
        # 빈줄 높이
        EMPTY_LINE = fs * ls + 2.0
        # 목차 항목: 작은 줄간격 (next=1pt)
        TOC_LINE = fs * ls + 1.0

        def _text_height(text):
            """단순 본문 텍스트의 높이(pt)"""
            if not text:
                return 0.0
            cpl = int(PAGE_W / fs)
            h = 0.0
            for ln in text.split('\n'):
                s = ln.strip()
                if not s:
                    h += EMPTY_LINE
                else:
                    h += max(1, -(-len(s) // cpl)) * (fs * ls) + 2.0
            return h

        def _extra_pages_pt(height_pt):
            """높이(pt) → 첫 페이지 이후 추가 페이지 수"""
            if height_pt <= PAGE_H:
                return 0
            return int((height_pt - 0.1) / PAGE_H)

        cur_page = 1  # 표지 = 1페이지

        # 목차 (page_break → 새 페이지)
        cur_page += 1
        toc_h = H1_HEIGHT + EMPTY_LINE + n_chapters * TOC_LINE + EMPTY_LINE
        cur_page += _extra_pages_pt(toc_h)

        # 가치 요약 (page_break → 새 페이지)
        analysis_data = ebook_data.get('analysis', {})
        if analysis_data:
            cur_page += 1
            val_h = H1_HEIGHT
            problem = analysis_data.get('problem_solved', {})
            for key in ('time', 'money', 'emotion'):
                val = problem.get(key, '')
                if val:
                    val_h += H2_HEIGHT + _text_height(val)
            if analysis_data.get('why_pay'):
                val_h += H2_HEIGHT + _text_height(analysis_data['why_pay'])
            cur_page += _extra_pages_pt(val_h)

        # 프롤로그 (page_break → 새 페이지)
        prologue_tmp = ebook_data.get('prologue', '')
        if prologue_tmp and prologue_tmp.strip():
            cur_page += 1
            pro_h = H1_HEIGHT + _text_height(prologue_tmp)
            cur_page += _extra_pages_pt(pro_h)

        # 각 챕터 (page_break → 새 페이지)
        chapter_est_pages = {}
        for ci, ch_data in enumerate(ebook_data.get('chapters_content', [])):
            ch_num = chapters[ci].get('chapter_num', ci + 1) if ci < n_chapters else ci + 1
            cur_page += 1  # 챕터 시작 (page_break=True)
            chapter_est_pages[ch_num] = cur_page
            content_tmp = ch_data.get('content', '')
            # 라벨(1줄) + H1 제목 + before + after + 빈줄 + 본문
            ch_h = BODY_LINE + H1_HEIGHT + BODY_LINE + BODY_LINE + EMPTY_LINE
            ch_h += self._calc_content_height_pt(
                content_tmp, fs, hs, ss, ls, PAGE_W)
            cur_page += _extra_pages_pt(ch_h)

        for ch in chapters:
            phase    = ch.get('phase', '')
            num      = ch.get('chapter_num', '')
            ch_title = ch.get('title', '')
            pg_num   = chapter_est_pages.get(num, '')
            # 점선 + 페이지번호 형태로 표시
            dots = '·' * max(2, 40 - len(f"[{phase}]  CHAPTER {num}  {ch_title}"))
            paras.append(self._toc(
                f"[{phase}]  CHAPTER {num}  {ch_title}  {dots} {pg_num}"
            ))
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

        # 프롤로그
        prologue = ebook_data.get('prologue', '')
        if prologue and prologue.strip():
            paras.append(self._h1('프롤로그', page_break=True))
            for line in prologue.split('\n'):
                stripped = line.strip()
                if stripped:
                    paras.append(self._body(stripped))
                else:
                    paras.append(self._empty())

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

        # 에필로그
        epilogue = ebook_data.get('epilogue', '')
        if epilogue and epilogue.strip():
            paras.append(self._h1('에필로그', page_break=True))
            for line in epilogue.split('\n'):
                stripped = line.strip()
                if stripped:
                    paras.append(self._body(stripped))
                else:
                    paras.append(self._empty())

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

        # ── XML 조립 (footer에서 자동 쪽번호 처리) ──────────────
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
