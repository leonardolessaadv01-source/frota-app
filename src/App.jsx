
9px",borderRadius:7,fontWeight:700,fontSize:11,cursor:"pointer",marginLeft:4}}>📥 Excel</button>
      </div>
    </div>

    <div style={{maxWidth:920,margin:"0 auto",padding:"20px 14px"}}>
      {section==="dashboard"   && <Dashboard vehicles={vehicles} manutencoes={manutencoes} contratos={contratos} reservas={reservas} onNav={goTo}/>}
      {section==="lista"       && <Frota vehicles={vehicles} fotos={fotos} onSave={crudSave(setVehicles)} onDelete={crudDel(setVehicles)} onFoto={handleFoto}/>}
      {section==="contratos"   && <Contratos vehicles={vehicles} clientes={clientes} contratos={contratos} fotos={fotos} onSave={crudSave(setContratos)} onDelete={crudDel(setContratos)} onFoto={handleFoto} showToast={showToast}/>}
      {section==="manutencoes" && <Manutencoes vehicles={vehicles} manutencoes={manutencoes} onSave={crudSave(setManut)} onDelete={crudDel(setManut)} showToast={showToast}/>}
      {section==="clientes"    && <Clientes clientes={clientes} contratos={contratos} onSave={crudSave(setClientes)} onDelete={crudDel(setClientes)} showToast={showToast}/>}
      {section==="documentos"  && <Documentos vehicles={vehicles} docs={docs} onSave={crudSave(setDocs)} onDelete={crudDel(setDocs)} showToast={showToast}/>}
      {section==="financeiro"  && <Financeiro vehicles={vehicles} manutencoes={manutencoes} contratos={contratos}/>}
      {section==="calendario"  && <Calendario vehicles={vehicles} contratos={contratos} reservas={reservas} onNav={goTo}/>}
      {section==="reservas"    && <Reservas vehicles={vehicles} reservas={reservas} clientes={clientes} onSave={crudSave(setReservas)} onDelete={crudDel(setReservas)} showToast={showToast}/>}
    </div>

    <Toast msg={toast.msg} color={toast.color}/>
  </div>;
}
