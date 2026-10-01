// Dados de EXEMPLO (fictícios) só para testar cálculos e o preview local
const hoje = '2026-10-01';
const P = (data, cliente, tipo, o) => Object.assign({data, cliente, tipo, postas:0,c200:0,c300:0,d200:0,d300:0,un:0,kg:0,entrega:'',valor:0,custo:0,margemR:0,margemPct:0}, o);
const mon = {
  clientes: [
    {nome:'Vegsim - Mooca', tipo:'Ativo', cidade:'São Paulo/SP', ciclo:14, antec:3, obs:'', motivo:'', registro:'2026-08-10'},
    {nome:'Vegsim Moema', tipo:'Ativo', cidade:'São Paulo/SP', ciclo:21, antec:3, obs:'', motivo:'', registro:'2026-08-20'},
    {nome:'Quatro Estrelas', tipo:'Ativo', cidade:'Campinas/SP', ciclo:21, antec:3, obs:'', motivo:'', registro:'2026-07-01'},
    {nome:'Manjerição do Quintal', tipo:'Ativo', cidade:'Piracicaba/SP', ciclo:30, antec:5, obs:'', motivo:'Item no cardápio, ninguém pediu', registro:'2026-06-15'},
    {nome:'Restô Exemplo Jardins', tipo:'Prospecção', cidade:'São Paulo/SP', ciclo:7, antec:0, obs:'', motivo:'', registro:'2026-09-10'},
    {nome:'Empório Exemplo', tipo:'Prospecção', cidade:'São Paulo/SP', ciclo:7, antec:0, obs:'', motivo:'', registro:'2026-09-25'},
    {nome:'Bistrô Convertido', tipo:'Prospecção', cidade:'São Paulo/SP', ciclo:7, antec:0, obs:'', motivo:'', registro:'2026-08-01'}
  ],
  pedidos: [
    P('2026-08-12','Vegsim - Mooca','Pedido',{postas:20,c200:10,un:30,kg:7,valor:1500,custo:560,margemR:940,margemPct:62}),
    P('2026-08-26','Vegsim - Mooca','Pedido',{postas:24,c200:10,un:34,kg:7.6,valor:1700,custo:620,margemR:1080,margemPct:63}),
    P('2026-09-10','Vegsim - Mooca','Pedido',{postas:30,c200:12,un:42,kg:9,valor:2050,custo:760,margemR:1290,margemPct:63}),
    P('2026-09-24','Vegsim - Mooca','Pedido',{postas:30,c200:12,un:42,kg:9,valor:2050,custo:760,margemR:1290,margemPct:63}),
    P('2026-08-22','Vegsim Moema','Pedido',{postas:12,d200:10,un:22,kg:4.6,valor:1000,custo:400,margemR:600,margemPct:60}),
    P('2026-09-05','Vegsim Moema','Pedido',{postas:12,d200:10,un:22,kg:4.6,valor:1000,custo:400,margemR:600,margemPct:60}),
    P('2026-07-03','Quatro Estrelas','Pedido',{c300:20,un:20,kg:6,valor:1100,custo:330,margemR:770,margemPct:70}),
    P('2026-07-25','Quatro Estrelas','Pedido',{c300:20,un:20,kg:6,valor:1100,custo:330,margemR:770,margemPct:70}),
    P('2026-08-14','Quatro Estrelas','Pedido',{c300:12,un:12,kg:3.6,valor:700,custo:200,margemR:500,margemPct:71}),
    P('2026-06-20','Manjerição do Quintal','Pedido',{postas:30,un:30,kg:6,valor:1500,custo:560,margemR:940,margemPct:62}),
    P('2026-07-22','Manjerição do Quintal','Pedido',{postas:20,un:20,kg:4,valor:1000,custo:375,margemR:625,margemPct:62}),
    P('2026-09-01','Restô Exemplo Jardins','Amostra',{postas:2,un:2,kg:.4,entrega:'2026-09-03'}),
    P('2026-09-26','Empório Exemplo','Amostra',{d200:2,un:2,kg:.4,entrega:''}),
    P('2026-08-05','Bistrô Convertido','Amostra',{postas:2,un:2,kg:.4,entrega:'2026-08-07'}),
    P('2026-08-20','Bistrô Convertido','Pedido',{postas:10,un:10,kg:2,valor:520,custo:190,margemR:330,margemPct:63})
  ],
  fups: [
    {data:'2026-09-12',cliente:'Manjerição do Quintal',canal:'WhatsApp',resultado:'🔇 Não respondeu',prox:'',obs:'',tipo:'Recompra'},
    {data:'2026-09-20',cliente:'Manjerição do Quintal',canal:'WhatsApp',resultado:'⏸️ Adiou',prox:'',obs:'',tipo:'Recompra'},
    {data:'2026-09-10',cliente:'Restô Exemplo Jardins',canal:'WhatsApp',resultado:'🔇 Não respondeu',prox:'',obs:'',tipo:'Amostra'},
    {data:'2026-09-28',cliente:'Quatro Estrelas',canal:'WhatsApp',resultado:'📅 Vai comprar',prox:'',obs:'',tipo:'Recompra'}
  ],
  sellout: [
    {data:'2026-09-15',cliente:'Vegsim - Mooca',un:30,obs:'',autor:'Isabelly'},
    {data:'2026-09-28',cliente:'Vegsim - Mooca',un:32,obs:'',autor:'Isabelly'},
    {data:'2026-09-20',cliente:'Manjerição do Quintal',un:2,obs:'quase ninguém pediu',autor:'Isabelly'},
    {data:'2026-09-25',cliente:'Vegsim Moema',un:20,obs:'',autor:'Isabelly'}
  ],
  fichas: { }
};
module.exports = { mon, hoje };
