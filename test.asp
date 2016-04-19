<HTML>
<BODY>

<%

public lOld
public m_bytes(9999999) 
public loutput
dim lCR

set lCR= server.createobject("StringHandle.clsTradeRebate")


lCR.LoadData "01", "01", "2011", "30", "03", "2011", "ALL", "","888"

lcr.FindTrade "20592709"

Response.write   lCR.ClientRebateFC & "," & lCR.AECommFC & "," & lCR.BrokerCommFC & "," & lCR.BrokerRebateFC & "," & lCR.IntroducerRebateFC & "," & lCR.ResearchCreditFC & "," & lCR.ResearchCreditCCY & "," & lCR.IntroducerRebateFC

set lstr = nothing



%>


</HTML>
</BODY>
