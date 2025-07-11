# DrawingArcIssue

This sample illustrates the issue we are having trying to calculate the end coordinates of the arcTo element in the presetShapeDefinition.xml.
It is written in c# and generates the cloud as a svg.
This sample uses the cloud shape definition's pathLst:
```xml
    <pathLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
      <path w="43200" h="43200">
        <moveTo>
          <pt x="3900" y="14370" />
        </moveTo>
        <arcTo wR="6753" hR="9190" stAng="-11429249" swAng="7426832" />
        <arcTo wR="5333" hR="7267" stAng="-8646143" swAng="5396714" />
        <arcTo wR="4365" hR="5945" stAng="-8748475" swAng="5983381" />
        <arcTo wR="4857" hR="6595" stAng="-7859164" swAng="7034504" />
        <arcTo wR="5333" hR="7273" stAng="-4722533" swAng="6541615" />
        <arcTo wR="6775" hR="9220" stAng="-2776035" swAng="7816140" />
        <arcTo wR="5785" hR="7867" stAng="37501" swAng="6842000" />
        <arcTo wR="6752" hR="9215" stAng="1347096" swAng="6910353" />
        <arcTo wR="7720" hR="10543" stAng="3974558" swAng="4542661" />
        <arcTo wR="4360" hR="5918" stAng="-16496525" swAng="8804134" />
        <arcTo wR="4345" hR="5945" stAng="-14809710" swAng="9151131" />
        <close />
      </path>
      <path w="43200" h="43200" fill="none" extrusionOk="false">
        <moveTo>
          <pt x="4693" y="26177" />
        </moveTo>
        <arcTo wR="4345" hR="5945" stAng="5204520" swAng="1585770" />
        <moveTo>
          <pt x="6928" y="34899" />
        </moveTo>
        <arcTo wR="4360" hR="5918" stAng="4416628" swAng="686848" />
        <moveTo>
          <pt x="16478" y="39090" />
        </moveTo>
        <arcTo wR="6752" hR="9215" stAng="8257449" swAng="844866" />
        <moveTo>
          <pt x="28827" y="34751" />
        </moveTo>
        <arcTo wR="6752" hR="9215" stAng="387196" swAng="959901" />
        <moveTo>
          <pt x="34129" y="22954" />
        </moveTo>
        <arcTo wR="5785" hR="7867" stAng="-4217541" swAng="4255042" />
        <moveTo>
          <pt x="41798" y="15354" />
        </moveTo>
        <arcTo wR="5333" hR="7273" stAng="1819082" swAng="1665090" />
        <moveTo>
          <pt x="38324" y="5426" />
        </moveTo>
        <arcTo wR="4857" hR="6595" stAng="-824660" swAng="891534" />
        <moveTo>
          <pt x="29078" y="3952" />
        </moveTo>
        <arcTo wR="4857" hR="6595" stAng="-8950887" swAng="1091722" />
        <moveTo>
          <pt x="22141" y="4720" />
        </moveTo>
        <arcTo wR="4365" hR="5945" stAng="-9809656" swAng="1061181" />
        <moveTo>
          <pt x="14000" y="5192" />
        </moveTo>
        <arcTo wR="6753" hR="9190" stAng="-4002417" swAng="739161" />
        <moveTo>
          <pt x="4127" y="15789" />
        </moveTo>
        <arcTo wR="6753" hR="9190" stAng="9459261" swAng="711490" />
      </path>
    </pathLst>
```
