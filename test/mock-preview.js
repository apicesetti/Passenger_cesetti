window.geotab = { addin: {} };
(function(){
  const now = Date.now(), iso = m => new Date(now - m*60000).toISOString();
  const b64 = h => btoa('RUID' + String.fromCharCode(0,0,0) + h.match(/../g).map(x=>String.fromCharCode(parseInt(x,16))).join(''));
  const pas = [['8C9A8C13','Ana Gómez'],['31012C0F','Luis Medina'],['A1B2C3D4','Raúl Sosa'],['11223344','Carla Funes'],['55667788','Diego Bustos'],['99AABBCC','Mario Vera'],['DDEEFF00','Pablo Ortiz'],['12345678','Elena Luna']];
  const store = [{id:'v1',details:{tipo:'vehiculo',deviceId:'bBA',clase:'combi',asientos:15}},{id:'v2',details:{tipo:'vehiculo',deviceId:'bBB',clase:'colectivo',asientos:45}},{id:'v3',details:{tipo:'vehiculo',deviceId:'bBC',clase:'camioneta',asientos:4}},{id:'pe1',details:{tipo:'persona',userId:'b1C8',rol:'pasajero',habilitado:false,motivo:'Inducción vencida',vence:'',vehiculos:[]}}]
    .concat(pas.map((p,i)=>({id:'p'+i,details:{tipo:'pasajero',hex:p[0],nombre:p[1],legajo:String(1000+i),empresa:'Tecpetrol',turno:'Mañana',activo:true}})));
  let k=0; const feed=[{id:'c'+k++,device:{id:'bBA'},dateTime:iso(80),data:'UkVBREVSIENPTk5FQ1RFRA=='}];
  pas.forEach((p,i)=>feed.push({id:'c'+k++,device:{id:'bBA'},dateTime:iso(40-i*2),data:b64(p[0])}));
  feed.push({id:'c'+k++,device:{id:'bBA'},dateTime:iso(20),data:b64('8C9A8C13')});
  feed.push({id:'c'+k++,device:{id:'bBA'},dateTime:iso(12),data:b64('0BADCAFE')});
  for(let i=0;i<31;i++) feed.push({id:'c'+k++,device:{id:'bBB'},dateTime:iso(30-i/2),data:b64((0x10000000+i*7919).toString(16).toUpperCase())});
  feed.push({id:'c'+k++,device:{id:'bBC'},dateTime:iso(15),data:b64('12345678')});
  const h=(m,p)=>{
    if(m==='Get'&&p.typeName==='AddInData')return store;
    if(m==='Get'&&p.typeName==='Device')return[{id:'bBA',name:'CESETTI_TEST',licensePlate:'AF123CD'},{id:'bBB',name:'COLECTIVO 07',licensePlate:'AE456GH'},{id:'bBC',name:'HILUX 12',licensePlate:'AG789JK'}];
    if(m==='GetFeed')return p.fromVersion?{data:[],toVersion:'2'}:{data:feed,toVersion:'1'};
    if(m==='Get'&&p.typeName==='DeviceStatusInfo')return[{device:{id:'bBA'},driver:{id:'u1'}},{device:{id:'bBB'},driver:{id:'u2'}}];
    if(m==='Get'&&p.typeName==='User'&&p.search.isDriver)return[{id:'b1C8',name:'PASAJERO_1',firstName:'PASAJERO',lastName:'1',keys:[{driverKeyType:'Custom',serialNumber:'0BADCAFE'}]},{id:'u1',name:'jperez',firstName:'Juan',lastName:'Pérez',keys:[{driverKeyType:'Nfc',serialNumber:'047AEB8A777A80'}]},{id:'u2',name:'mrios',firstName:'Marta',lastName:'Ríos',keys:[]}];
    if(m==='Get'&&p.typeName==='CustomData')return feed.filter(f=>f.device.id===p.search.deviceSearch.id);
    if(m==='Get'&&p.typeName==='LogRecord'){
      const f=Date.parse(p.search.fromDate),tt=Date.parse(p.search.toDate);
      const pt=t=>{const k=Math.max(0,Math.min(1,(t-(now-90*60000))/(90*60000)));return{latitude:-38.9504-0.25*k+0.03*Math.sin(k*9),longitude:-68.0576+0.35*k,dateTime:new Date(t).toISOString()}};
      if(f===tt)return[pt(f)];
      const out=[];for(let t=Math.max(f,now-90*60000);t<=tt;t+=30000)out.push(pt(t));return out;}
    if(m==='Get'&&p.typeName==='IoxAddOn')return[{type:4247,device:{id:'bBA'},dateTime:'2026-11-12T16:55:40.918Z'},{type:4247,device:{id:'bBB'},dateTime:'2026-11-12T16:55:40.918Z'}];
    if(m==='Get'&&p.typeName==='AddInData')return store;
    if(m==='GetAddresses')return p.coordinates.map((c,i)=>({formattedAddress:'Ruta Provincial 7 km '+(10+i)+', Añelo, Neuquén',zones:i%3?[]:[{name:'Base Fortín de Piedra'}]}));
    if(m==='Get'&&p.typeName==='User')return[{u1:{firstName:'Juan',lastName:'Pérez'},u2:{firstName:'Marta',lastName:'Ríos'}}[p.search.id]];
    return [];
  };
  window.__api={call(m,p,ok){setTimeout(()=>ok(h(m,p)),5)},multiCall(cs,ok){setTimeout(()=>ok(cs.map(c=>h(c[0],c[1]))),5)}};
})();
