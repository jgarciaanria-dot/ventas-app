const Contabilidad = {
  async getCatalogoServicios() {
    try {
      const url = CONFIG.contabilidad.scriptUrl + '?action=getCatalogoServicios&_=' + Date.now();
      const res = await fetch(url);
      const data = await res.json();
      return data.servicios || [];
    } catch (e) {
      console.error('Error cargando catalogo:', e);
      return [];
    }
  },

  async getColaboradores() {
    try {
      const url = CONFIG.contabilidad.scriptUrl + '?action=getColaboradores&_=' + Date.now();
      const res = await fetch(url);
      const data = await res.json();
      return data.colaboradores || [];
    } catch (e) {
      console.error('Error cargando colaboradores:', e);
      return [];
    }
  },

  async getDashboardResumen(filtros) {
    try {
      const params = new URLSearchParams({ action: 'getDashboardResumen', _: Date.now() });
      if (filtros.inicio) params.append('inicio', filtros.inicio);
      if (filtros.fin) params.append('fin', filtros.fin);
      if (filtros.colaborador) params.append('colaborador', filtros.colaborador);
      const url = CONFIG.contabilidad.scriptUrl + '?' + params.toString();
      const res = await fetch(url);
      const data = await res.json();
      return data.resumen || { ventas:0, comision:0, studio:0, servicios:0, adelantos:0, neto:0, pagos:{} };
    } catch (e) {
      console.error('Error cargando resumen:', e);
      return { ventas:0, comision:0, studio:0, servicios:0, adelantos:0, neto:0, pagos:{} };
    }
  },

  async getDetalleServicios(filtros) {
    try {
      const params = new URLSearchParams({ action: 'getDetalleServicios', _: Date.now() });
      if (filtros.inicio) params.append('inicio', filtros.inicio);
      if (filtros.fin) params.append('fin', filtros.fin);
      if (filtros.colaborador) params.append('colaborador', filtros.colaborador);
      const url = CONFIG.contabilidad.scriptUrl + '?' + params.toString();
      const res = await fetch(url);
      const data = await res.json();
      return data.detalle || [];
    } catch (e) {
      console.error('Error cargando detalle:', e);
      return [];
    }
  },

  async getComisionesPorColaborador() {
    try {
      const url = CONFIG.contabilidad.scriptUrl + '?action=getComisionesPorColaborador&_=' + Date.now();
      const res = await fetch(url);
      const data = await res.json();
      return data.comisiones || {};
    } catch (e) {
      console.error('Error cargando comisiones:', e);
      return {};
    }
  },

  async registrarVenta(venta) {
    const res = await fetch(CONFIG.contabilidad.scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'registrarVenta', ...venta })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Error al registrar venta');
    return data.venta;
  },

  async registrarAdelanto(adelanto) {
    const res = await fetch(CONFIG.contabilidad.scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'registrarAdelanto', ...adelanto })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Error al registrar adelanto');
    return data;
  },

  async registrarPropina(propina) {
    const res = await fetch(CONFIG.contabilidad.scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'registrarPropina', ...propina })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Error al registrar propina');
    return data;
  },

  async getPropinasPendientes(colaborador) {
    try {
      const url = CONFIG.contabilidad.scriptUrl + '?action=getPropinasPendientes&colaborador=' + encodeURIComponent(colaborador) + '&_=' + Date.now();
      const res = await fetch(url);
      const data = await res.json();
      return data.pendientes || [];
    } catch (e) {
      console.error('Error cargando propinas pendientes:', e);
      return [];
    }
  },

  async marcarPropinaPagada(datos) {
    const res = await fetch(CONFIG.contabilidad.scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'marcarPropinaPagada', ...datos })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Error al marcar como pagada');
    return data;
  },

  async generarComprobante(datos) {
    const res = await fetch(CONFIG.contabilidad.scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'generarComprobante', ...datos })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Error al generar comprobante');
    return data;
  },

  async cerrarQuincena(datos) {
    const res = await fetch(CONFIG.contabilidad.scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'cerrarQuincena', ...datos })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Error al cerrar quincena');
    return data;
  }
};
