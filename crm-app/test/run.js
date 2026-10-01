const fs=require('fs'), vm=require('vm');
const m=require('../Inteligencia.js'); const {mon,hoje}=require('./sample.js');
const c=m.calcCarteira_(mon,hoje);
console.log('RESUMO', JSON.stringify(c.resumo));
c.ativos.forEach(a=>console.log(a.nome.padEnd(24), String(a.score).padStart(3), a.nivel.padEnd(8), 'dias',a.diasDesde,'ciclo',a.ciclo,'prox',a.proxRecompra,'|',a.flags.join('; ')));
c.prospects.forEach(p=>console.log('PROSPECT',p.nome.padEnd(24),p.status,p.diasDesde,p.diasParaEntregar));
const s=m.calcSell_(mon,hoje); s.linhas.forEach(l=>console.log('SELL',l.nome.padEnd(24),l.alerta.padEnd(13),'in90',l.sellIn.d90,'out90',l.sellOut.d90,'giro',l.giro&&l.giro.toFixed(2),'cob',l.coberturaSemanas));
const p=m.calcPrevisaoBase_(mon,hoje,c); console.log('PREV',JSON.stringify({t:p.ticketNovo,kg:p.kgNovo,rec:p.recompra30.valor,kgr:p.recompra30.kg,n:p.recompra30.clientes.length, tm:p.ticketMensalPorCliente}));
