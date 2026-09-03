'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  ArrowDownToLine,
  Bell,
  CalendarDays,
  Check,
  ChevronDown,
  ChevronRight,
  CircleDollarSign,
  ClipboardList,
  CreditCard,
  FileText,
  HandCoins,
  Home,
  Menu,
  MoreHorizontal,
  Search,
  UserRound,
  Users,
  WalletCards,
  X,
} from 'lucide-react'

import { Button } from '@/components/ui/button'

type Invoice = {
  id: string
  number: string
  client: string
  taxId: string
  due: string
  amount: number
  paid: number
}

type Payment = {
  method: string
  amount: number
  issueDate?: string
  paymentDate?: string
  issuingBank?: string
  bankAccount?: string
  issuerTaxId?: string
  eCheq?: 'Sí' | 'No'
  concept?: string
  receiptName?: string
}

const initialInvoices: Invoice[] = [
  { id: '1', number: 'FC A 0001-00051377', client: 'Blossom San Fernando (Diocla SRL)', taxId: '30-71683193-7', due: '05/09/2026', amount: 571707, paid: 0 },
  { id: '2', number: 'FC B 0001-00051401', client: 'Academia Nacional de Bellas Artes', taxId: '30-71683193-7', due: '28/08/2026', amount: 373650, paid: 0 },
  { id: '3', number: 'FC A 0001-00051215', client: 'Aguilera Verón Paul Nicolás', taxId: '20-32156480-9', due: '18/08/2026', amount: 56250, paid: 0 },
  { id: '4', number: 'FC A 0001-00051095', client: 'The Coffee Store Quilmes TCS', taxId: '30-70928411-2', due: '13/08/2026', amount: 621200, paid: 0 },
  { id: '5', number: 'FC B 0001-00051175', client: 'Heladería Trevi San Miguel', taxId: '30-71132041-8', due: '16/08/2026', amount: 215600, paid: 0 },
  { id: '6', number: 'FC A 0001-00051263', client: 'Volta Coffee Store Banfield', taxId: '30-72889012-5', due: '20/08/2026', amount: 394960, paid: 0 },
]

const money = (value: number) => new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', maximumFractionDigits: 2 }).format(value)
const methods = ['Depósito bancario', 'Efectivo', 'Transferencia', 'Cheque', 'Tarjeta de crédito']
const bankAccounts = ['Banco Nación · Cta. Cte. 001-234567/8', 'Banco Galicia · Cta. Cte. 104-778901/2', 'Santander · Cta. Cte. 655-009812/4']

export default function Page() {
  const [invoices, setInvoices] = useState(initialInvoices)
  const [selected, setSelected] = useState<string[]>(['1', '2'])
  const [query, setQuery] = useState('')
  const [tab, setTab] = useState('Pendientes')
  const [showPayment, setShowPayment] = useState(false)
  const [payments, setPayments] = useState<Payment[]>([{ method: 'Transferencia', amount: 610000 }])
  const [client, setClient] = useState('Todos los clientes')
  const [paymentGroup, setPaymentGroup] = useState('Todos los grupos')
  const [dateFrom, setDateFrom] = useState('2026-08-01')
  const [dateTo, setDateTo] = useState('2026-09-30')
  const [saved, setSaved] = useState(false)

  const visible = useMemo(() => invoices.filter((invoice) => {
    const matchesQuery = `${invoice.client} ${invoice.number}`.toLowerCase().includes(query.toLowerCase())
    const balance = invoice.amount - invoice.paid
    const matchesTab = tab !== 'Pagos remanentes' && (tab === 'Todos' || (tab === 'Pendientes' && balance > 0) || (tab === 'Morosos' && balance > 0 && invoice.due < '02/09/2026'))
    return matchesQuery && matchesTab
  }), [invoices, query, tab])
  const totalDebt = invoices.reduce((sum, invoice) => sum + invoice.amount - invoice.paid, 0)
  const selectedInvoices = invoices.filter((invoice) => selected.includes(invoice.id))
  const applied = Math.min(payments.reduce((sum, payment) => sum + payment.amount, 0), selectedInvoices.reduce((sum, invoice) => sum + invoice.amount - invoice.paid, 0))
  const paidIn = payments.reduce((sum, payment) => sum + payment.amount, 0)
  const remainder = Math.max(paidIn - applied, 0)

  function toggleInvoice(id: string) {
    setSelected((current) => current.includes(id) ? current.filter((item) => item !== id) : [...current, id])
  }

  function addPayment() { setPayments((current) => [...current, { method: 'Efectivo', amount: 0 }]) }
  function updatePayment(index: number, changes: Partial<Payment>) {
    setPayments((current) => current.map((item, i) => i === index ? { ...item, ...changes } : item))
  }
  function exportCollected() {
    const collected = invoices.filter((invoice) => {
      const [day, month, year] = invoice.due.split('/')
      const paidDate = `${year}-${month}-${day}`
      return invoice.paid > 0 && paidDate >= dateFrom && paidDate <= dateTo
    })
    const rows = collected.map((invoice) => ({
      Cliente: invoice.client,
      CUIT: invoice.taxId,
      Comprobante: invoice.number,
      Fecha: invoice.due,
      Importe: invoice.amount,
      Cobrado: invoice.paid,
      Saldo: invoice.amount - invoice.paid,
    }))
    const worksheet = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Mensaje: 'No hay cobros en el periodo seleccionado' }])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Cobros')
    XLSX.writeFile(workbook, `cobros-${dateFrom}-${dateTo}.xlsx`)
  }

  function savePayment() {
    let remaining = paidIn
    setInvoices((current) => current.map((invoice) => {
      if (!selected.includes(invoice.id) || remaining <= 0) return invoice
      const balance = invoice.amount - invoice.paid
      const allocation = Math.min(balance, remaining)
      remaining -= allocation
      return { ...invoice, paid: invoice.paid + allocation }
    }))
    setSaved(true)
    setTimeout(() => { setShowPayment(false); setSaved(false) }, 1100)
  }

  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="topbar navbar d-flex align-items-center justify-content-between px-4">
        <div className="flex items-center gap-5"><span className="text-xl font-semibold tracking-tight">Administración</span><Menu className="size-6" /></div>
        <div className="flex items-center gap-6"><button className="company-select">DEPCSUITE SA <ChevronDown className="size-4" /></button><Bell className="size-5 text-muted-foreground" /><div className="hidden items-center gap-2 text-sm font-medium text-muted-foreground md:flex"><UserRound className="size-5" /> Nelson Daniel Tarche</div></div>
      </header>
      <div className="container-fluid d-flex min-vh-100 p-0">
        <aside className="sidebar d-none d-lg-flex flex-column gap-2 p-4 col-lg-2">
          <div className="mb-5 px-3 text-xs font-semibold uppercase tracking-[0.18em] text-sidebar-muted">Módulos</div>
          {[[Home, 'Inicio'], [ClipboardList, 'Dashboard'], [FileText, 'Pedidos'], [CircleDollarSign, 'Productos'], [Users, 'Clientes'], [WalletCards, 'Ventas'], [CreditCard, 'Pagos'], [HandCoins, 'Cobranza'], [FileText, 'Cheques'], [Home, 'Contratos']].map(([Icon, label]) => <button key={label as string} className={`side-item ${label === 'Cobranza' ? 'active' : ''}`}><Icon className="size-[18px]" /> <span>{label as string}</span>{label !== 'Inicio' && <ChevronRight className="ml-auto size-4 opacity-60" />}</button>)}
        </aside>
        <section className="col flex-grow-1 overflow-hidden">
          <div className="mx-auto max-w-[1500px] p-5 md:p-8">
            <div className="mb-7 flex flex-wrap items-end justify-between gap-4"><div><div className="mb-2 flex items-center gap-2 text-sm text-muted-foreground"><span>Inicio</span><ChevronRight className="size-4" /><span className="text-primary">Cobranza</span></div><h1 className="text-3xl font-semibold tracking-tight">Cobranza</h1><p className="mt-1 text-sm text-muted-foreground">Gestioná pagos, imputaciones y saldos a favor de tus clientes.</p></div><Button onClick={() => setShowPayment(true)} className="h-11 gap-2"><HandCoins className="size-4" data-icon="inline-start" /> Registrar cobro</Button></div>
            <div className="summary-grid mb-6"><Summary label="Pendiente de cobro" value={money(totalDebt)} detail="6 facturas abiertas" tone="navy" icon={CircleDollarSign} /><Summary label="Cobrado este mes" value={money(1500000)} detail="+12,4% vs. mes anterior" tone="blue" icon={Check} /><Summary label="Saldos a favor" value={money(6100)} detail="3 remanentes activos" tone="gold" icon={WalletCards} /><Summary label="Clientes morosos" value="8" detail="Requieren seguimiento" tone="red" icon={Users} /></div>
            <div className="filter-panel mb-6"><div className="flex flex-wrap items-center gap-3"><div className="search-wrap"><Search className="size-4 text-muted-foreground" /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Buscar cliente o factura..." /></div><select value={client} onChange={(event) => setClient(event.target.value)} className="filter-select form-select"><option>Todos los clientes</option><option>Blossom San Fernando</option><option>The Coffee Store Quilmes</option></select><label className="date-filter"><CalendarDays className="size-4" /><span>Desde</span><input aria-label="Fecha desde" type="date" value={dateFrom} onChange={(event) => setDateFrom(event.target.value)} /><span>Hasta</span><input aria-label="Fecha hasta" type="date" value={dateTo} onChange={(event) => setDateTo(event.target.value)} /></label><Button variant="outline" onClick={exportCollected} className="gap-2"><ArrowDownToLine className="size-4" data-icon="inline-start" /> Excel</Button></div></div>
            <div className="mb-4 flex items-center justify-between"><div className="tab-list">{['Todos', 'Pendientes', 'Morosos', 'Pagos remanentes'].map((item) => <button key={item} onClick={() => setTab(item)} className={tab === item ? 'selected' : ''}>{item}<span>{item === 'Todos' ? invoices.length : item === 'Pendientes' ? invoices.filter((invoice) => invoice.amount > invoice.paid).length : 4}</span></button>)}</div><span className="text-sm text-muted-foreground">Mostrando {visible.length} de {invoices.length} comprobantes</span></div>
            <div className="table-card"><div className="table-scroll"><table><thead><tr><th className="w-10"><input type="checkbox" checked={selected.length === visible.length && visible.length > 0} onChange={() => setSelected(selected.length === visible.length ? [] : visible.map((invoice) => invoice.id))} aria-label="Seleccionar todas" /></th><th>Cliente</th><th>Comprobante</th><th>Vencimiento</th><th>Estado</th><th className="text-right">Importe</th><th className="text-right">Saldo</th><th /></tr></thead><tbody>{visible.map((invoice) => { const balance = invoice.amount - invoice.paid; const overdue = invoice.due < '02/09/2026'; return <tr key={invoice.id}><td><input type="checkbox" checked={selected.includes(invoice.id)} onChange={() => toggleInvoice(invoice.id)} aria-label={`Seleccionar ${invoice.number}`} /></td><td><div className="font-medium">{invoice.client}</div><div className="text-xs text-muted-foreground">CUIT {invoice.taxId}</div></td><td><span className="document">{invoice.number}</span></td><td>{invoice.due}<div className="text-xs text-muted-foreground">{overdue ? 'Vencida' : 'Por vencer'}</div></td><td><span className={`status ${overdue ? 'danger' : 'warning'}`}>{overdue ? 'Morosa' : 'Pendiente'}</span></td><td className="text-right font-medium">{money(invoice.amount)}</td><td className="text-right font-semibold text-danger">{money(balance)}</td><td><button className="icon-button" aria-label="Más opciones"><MoreHorizontal className="size-4" /></button></td></tr>})}</tbody></table></div><div className="table-footer"><span><strong>{selected.length}</strong> seleccionadas</span><span className="font-medium">Saldo seleccionado: <strong>{money(selectedInvoices.reduce((sum, invoice) => sum + invoice.amount - invoice.paid, 0))}</strong></span></div></div>
          </div>
        </section>
      </div>
      {showPayment && <div className="modal-backdrop"><div className="payment-modal"><div className="modal-header"><div><p className="eyebrow">Nuevo movimiento</p><h2>Registrar cobro</h2><p>Imputá uno o varios pagos a las facturas del cliente.</p></div><button onClick={() => setShowPayment(false)} className="icon-button"><X className="size-5" /></button></div><div className="modal-body"><label className="field-label">Grupo<select value={paymentGroup} onChange={(event) => { setPaymentGroup(event.target.value); setClient('Todos los clientes') }}><option>Todos los grupos</option><option>Grupo Blossom</option><option>Grupo Academia</option></select></label><label className="field-label">Cliente<select value={client} onChange={(event) => setClient(event.target.value)}><option>Todos los clientes</option>{(paymentGroup === 'Todos los grupos' || paymentGroup === 'Grupo Blossom') && <option>Blossom San Fernando (Diocla SRL)</option>}{(paymentGroup === 'Todos los grupos' || paymentGroup === 'Grupo Academia') && <option>Academia Nacional de Bellas Artes</option>}</select></label><div className="section-heading"><span>Medios de pago</span><button onClick={addPayment}>+ Agregar otro pago</button></div>{payments.map((payment, index) => <div className="payment-row" key={`${payment.method}-${index}`}><select value={payment.method} onChange={(event) => updatePayment(index, { method: event.target.value, bankAccount: event.target.value === 'Transferencia' ? payment.bankAccount : undefined })}>{methods.map((method) => <option key={method}>{method}</option>)}</select><div className="amount-input"><span>$</span><input type="number" value={payment.amount} onChange={(event) => setPayments((current) => current.map((item, i) => i === index ? { ...item, amount: Number(event.target.value) } : item))} /></div>{payment.method === 'Transferencia' && <select aria-label="Cuenta bancaria destino" value={payment.bankAccount ?? ''} onChange={(event) => updatePayment(index, { bankAccount: event.target.value })}><option value="">Seleccioná la cuenta bancaria</option>{bankAccounts.map((account) => <option key={account} value={account}>{account}</option>)}</select>}{payments.length > 1 && <button onClick={() => setPayments((current) => current.filter((_, i) => i !== index))} className="remove-payment"><X className="size-4" /></button>}{payment.method === 'Cheque' && <div className="cheque-details">
  <label className="field-label">Fecha de emisión<input type="date" value={payment.issueDate ?? ''} onChange={(event) => updatePayment(index, { issueDate: event.target.value })} /></label>
  <label className="field-label">Fecha de pago<input type="date" value={payment.paymentDate ?? ''} onChange={(event) => updatePayment(index, { paymentDate: event.target.value })} /></label>
  <label className="field-label">Banco emisor<input value={payment.issuingBank ?? ''} onChange={(event) => updatePayment(index, { issuingBank: event.target.value })} placeholder="Nombre del banco" /></label>
  <label className="field-label">CUIT emisor<input value={payment.issuerTaxId ?? ''} onChange={(event) => updatePayment(index, { issuerTaxId: event.target.value })} placeholder="00-00000000-0" /></label>
  <label className="field-label">E-Cheq<select value={payment.eCheq ?? 'No'} onChange={(event) => updatePayment(index, { eCheq: event.target.value as 'Sí' | 'No' })}><option>No</option><option>Sí</option></select></label>
  <label className="field-label cheque-concept">Concepto<input value={payment.concept ?? ''} onChange={(event) => updatePayment(index, { concept: event.target.value })} placeholder="Concepto del cheque" /></label>
  <label className="field-label cheque-upload">Adjuntar comprobante<input type="file" accept="image/*,.pdf" onChange={(event) => updatePayment(index, { receiptName: event.target.files?.[0]?.name ?? '' })} />{payment.receiptName && <span className="text-xs text-muted-foreground">{payment.receiptName}</span>}</label>
</div>}</div>)}<div className="section-heading mt-6"><span>Facturas a cancelar</span><button onClick={() => setSelected(visible.map((invoice) => invoice.id))}>Seleccionar todas</button></div><div className="invoice-picker">{selectedInvoices.length ? selectedInvoices.map((invoice) => <div className="picker-row" key={invoice.id}><input type="checkbox" checked onChange={() => toggleInvoice(invoice.id)} /><span>{invoice.number}</span><span className="ml-auto">{money(invoice.amount - invoice.paid)}</span></div>) : <p className="empty-picker">Seleccioná facturas desde la tabla para imputar el pago.</p>}</div><div className="allocation"><div><span>Pagado</span><strong>{money(paidIn)}</strong></div><div><span>Para cancelar</span><strong className="text-primary">{money(applied)}</strong></div><div><span>Remanente / saldo a favor</span><strong className={remainder ? 'text-success' : ''}>{money(remainder)}</strong></div></div></div><div className="modal-footer"><Button variant="outline" onClick={() => setShowPayment(false)}>Cancelar</Button><Button onClick={savePayment} disabled={!selectedInvoices.length || !paidIn}>{saved ? 'Cobro guardado' : 'Confirmar e imputar'}</Button></div></div></div>}
    </main>
  )
}

function Summary({ label, value, detail, tone, icon: Icon }: { label: string; value: string; detail: string; tone: string; icon: typeof CircleDollarSign }) { return <div className={`summary-card ${tone}`}><div className="summary-icon"><Icon className="size-5" /></div><div><p>{label}</p><strong>{value}</strong><span>{detail}</span></div></div> }
