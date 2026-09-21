import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import "./styles.css";

const initialCustomers = [
  { id: 1, firstName: "Anna", lastName: "Meier", email: "anna.meier@example.com", phone: "+41 79 123 45 67", city: "Zug", service: "Reinigung", status: "Aktiv", signup: "21.09.2026" },
  { id: 2, firstName: "Lukas", lastName: "Keller", email: "lukas.keller@example.com", phone: "+41 78 222 11 33", city: "Baar", service: "Unterhaltsreinigung", status: "Aktiv", signup: "19.09.2026" },
  { id: 3, firstName: "Sarah", lastName: "Müller", email: "sarah.mueller@example.com", phone: "+41 76 555 12 34", city: "Cham", service: "Fensterreinigung", status: "Inaktiv", signup: "15.09.2026" }
];

const initialCompanies = [
  { id: 1, company: "Beispiel Immobilien AG", contact: "Max Muster", email: "max@example.com", phone: "+41 41 000 00 00", industry: "Immobilien", interest: "Gebäudereinigung", status: "Interessiert" },
  { id: 2, company: "Zuger Office GmbH", contact: "Nina Keller", email: "nina@example.com", phone: "+41 79 111 22 33", industry: "Dienstleistungen", interest: "Büroreinigung", status: "Kontakt" }
];

function App() {
  const [page, setPage] = useState("dashboard");
  const [customers, setCustomers] = useState(initialCustomers);
  const [companies, setCompanies] = useState(initialCompanies);
  const [query, setQuery] = useState("");
  const [selectedCustomer, setSelectedCustomer] = useState(null);
  const [showAdd, setShowAdd] = useState(false);

  const filteredCustomers = useMemo(() => {
    const q = query.toLowerCase().trim();
    if (!q) return customers;
    return customers.filter(c =>
      Object.values(c).some(v => String(v).toLowerCase().includes(q))
    );
  }, [customers, query]);

  const addCustomer = (e) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    const customer = {
      id: Date.now(),
      firstName: f.get("firstName"),
      lastName: f.get("lastName"),
      email: f.get("email"),
      phone: f.get("phone"),
      city: f.get("city"),
      service: f.get("service"),
      status: "Aktiv",
      signup: new Date().toLocaleDateString("de-CH")
    };
    setCustomers([customer, ...customers]);
    setShowAdd(false);
  };

  const importExcel = async (file) => {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const imported = rows.map((r, i) => ({
      id: Date.now() + i,
      firstName: r["First Name"] || r["Vorname"] || r["firstName"] || "",
      lastName: r["Last Name"] || r["Nachname"] || r["lastName"] || "",
      email: r["Email"] || r["E-Mail"] || r["email"] || "",
      phone: r["Phone"] || r["Telefon"] || r["phone"] || "",
      city: r["City"] || r["Ort"] || r["city"] || "",
      service: r["Service"] || r["Dienstleistung"] || r["service"] || "",
      status: r["Status"] || "Aktiv",
      signup: r["Signup Date"] || r["Anmeldedatum"] || new Date().toLocaleDateString("de-CH")
    }));

    setCustomers(prev => {
      const byEmail = new Map(prev.map(c => [c.email.toLowerCase(), c]));
      imported.forEach(c => {
        if (c.email && byEmail.has(c.email.toLowerCase())) {
          byEmail.set(c.email.toLowerCase(), { ...byEmail.get(c.email.toLowerCase()), ...c });
        } else {
          byEmail.set((c.email || `id-${c.id}`).toLowerCase(), c);
        }
      });
      return [...byEmail.values()];
    });
    setPage("customers");
  };

  const exportExcel = () => {
    const ws = XLSX.utils.json_to_sheet(customers);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Customers");
    XLSX.writeFile(wb, "sponti-customers.xlsx");
  };

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="brand"><div className="logo">S</div><div><strong>SPONTI</strong><span>CRM</span></div></div>
        <nav>
          <button className={page==="dashboard"?"active":""} onClick={()=>setPage("dashboard")}>⌂ <span>Dashboard</span></button>
          <button className={page==="customers"?"active":""} onClick={()=>setPage("customers")}>♙ <span>Kunden</span></button>
          <button className={page==="companies"?"active":""} onClick={()=>setPage("companies")}>▦ <span>Firmen</span></button>
          <button className={page==="import"?"active":""} onClick={()=>setPage("import")}>↥ <span>Excel Import</span></button>
        </nav>
        <div className="side-bottom"><div className="avatar">S</div><div><b>Sponti Admin</b><small>Administrator</small></div></div>
      </aside>

      <main>
        <header>
          <div>
            <h1>{page==="dashboard"?"Dashboard":page==="customers"?"Kunden":page==="companies"?"Firmen":"Excel Import"}</h1>
            <p>Sponti Kundenverwaltung</p>
          </div>
          <div className="header-actions">
            <input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Suchen..." />
            <button className="primary" onClick={()=>setShowAdd(true)}>+ Neuer Kunde</button>
          </div>
        </header>

        {page === "dashboard" && <Dashboard customers={customers} companies={companies} setPage={setPage} />}
        {page === "customers" && <Customers customers={filteredCustomers} onSelect={setSelectedCustomer} onExport={exportExcel} />}
        {page === "companies" && <Companies companies={companies} />}
        {page === "import" && <ImportPage onImport={importExcel} />}

        {selectedCustomer && <CustomerModal customer={selectedCustomer} onClose={()=>setSelectedCustomer(null)} />}
        {showAdd && <AddModal onClose={()=>setShowAdd(false)} onSubmit={addCustomer} />}
      </main>
    </div>
  );
}

function Dashboard({customers, companies, setPage}) {
  const active = customers.filter(c=>c.status==="Aktiv").length;
  return <section>
    <div className="cards">
      <Card label="Kunden total" value={customers.length} icon="♙" />
      <Card label="Aktive Kunden" value={active} icon="✓" />
      <Card label="Firmen" value={companies.length} icon="▦" />
      <Card label="Neue Kunden" value={customers.length ? Math.min(customers.length, 12) : 0} icon="↗" />
    </div>
    <div className="grid">
      <div className="panel large">
        <div className="panel-head"><div><h2>Neueste Kunden</h2><p>Die zuletzt registrierten Kunden</p></div><button onClick={()=>setPage("customers")}>Alle anzeigen →</button></div>
        <CustomerTable customers={customers.slice(0,5)} onSelect={()=>{}} compact />
      </div>
      <div className="panel">
        <div className="panel-head"><div><h2>Firmen</h2><p>Potenzielle Geschäftspartner</p></div></div>
        {companies.map(c=><div className="company-row" key={c.id}><div className="company-icon">▦</div><div><b>{c.company}</b><small>{c.interest}</small></div><span className="pill">{c.status}</span></div>)}
      </div>
    </div>
  </section>
}

function Card({label,value,icon}) { return <div className="card"><div><small>{label}</small><strong>{value}</strong></div><div className="card-icon">{icon}</div></div> }

function Customers({customers,onSelect,onExport}) {
  return <section><div className="toolbar"><div><h2>Alle Kunden</h2><p>{customers.length} Kunden gefunden</p></div><button onClick={onExport}>↓ Excel exportieren</button></div><div className="panel"><CustomerTable customers={customers} onSelect={onSelect} /></div></section>
}

function CustomerTable({customers,onSelect,compact}) {
  return <div className="table-wrap"><table><thead><tr><th>Kunde</th><th>E-Mail</th><th>Telefon</th><th>Ort</th><th>Dienstleistung</th><th>Status</th><th></th></tr></thead><tbody>
    {customers.map(c=><tr key={c.id} onClick={()=>onSelect(c)}><td><div className="person"><div className="mini-avatar">{(c.firstName[0]||"")+(c.lastName[0]||"")}</div><div><b>{c.firstName} {c.lastName}</b><small>Seit {c.signup}</small></div></div></td><td>{c.email}</td><td>{c.phone}</td><td>{c.city}</td><td>{c.service}</td><td><span className={"status "+(c.status==="Aktiv"?"green":"gray")}>{c.status}</span></td><td>⋯</td></tr>)}
  </tbody></table>{!customers.length && <div className="empty">Keine Kunden gefunden.</div>}</div>
}

function Companies({companies}) {
  return <section><div className="toolbar"><div><h2>Firmen</h2><p>Unternehmen, die mit Sponti arbeiten möchten</p></div><button className="primary">+ Firma hinzufügen</button></div><div className="panel"><table><thead><tr><th>Firma</th><th>Kontakt</th><th>E-Mail</th><th>Branche</th><th>Interesse</th><th>Status</th></tr></thead><tbody>{companies.map(c=><tr key={c.id}><td><b>{c.company}</b></td><td>{c.contact}</td><td>{c.email}</td><td>{c.industry}</td><td>{c.interest}</td><td><span className="status">{c.status}</span></td></tr>)}</tbody></table></div></section>
}

function ImportPage({onImport}) {
  const [fileName,setFileName]=useState("");
  return <section><div className="import-card panel"><div className="upload-icon">↑</div><h2>Excel-Datei importieren</h2><p>Lade die Excel-Datei aus deinem bestehenden Login-System hoch. Sponti übernimmt die Kundendaten in die Kundenverwaltung.</p><label className="upload"><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>{setFileName(e.target.files[0]?.name||""); if(e.target.files[0]) onImport(e.target.files[0])}}/>Datei auswählen</label>{fileName&&<div className="file">✓ {fileName}</div>}<div className="mapping"><b>Unterstützte Spalten</b><span>Vorname · Nachname · E-Mail · Telefon · Ort · Dienstleistung · Status · Anmeldedatum</span></div></div></section>
}

function CustomerModal({customer,onClose}) { return <div className="overlay" onClick={onClose}><div className="modal" onClick={e=>e.stopPropagation()}><button className="close" onClick={onClose}>×</button><div className="modal-person"><div className="big-avatar">{customer.firstName[0]}{customer.lastName[0]}</div><div><h2>{customer.firstName} {customer.lastName}</h2><p>Kunde seit {customer.signup}</p></div></div><div className="details">{Object.entries({E-Mail:customer.email,Telefon:customer.phone,Ort:customer.city,Dienstleistung:customer.service,Status:customer.status}).map(([k,v])=><div key={k}><small>{k}</small><b>{v}</b></div>)}</div><h3>Notizen</h3><textarea placeholder="Notizen zum Kunden..."></textarea><button className="primary full" onClick={onClose}>Schliessen</button></div></div> }

function AddModal({onClose,onSubmit}) { return <div className="overlay" onClick={onClose}><form className="modal" onClick={e=>e.stopPropagation()} onSubmit={onSubmit}><button type="button" className="close" onClick={onClose}>×</button><h2>Neuen Kunden hinzufügen</h2><div className="form-grid">{["firstName","lastName","email","phone","city","service"].map((x,i)=><label key={x}>{["Vorname","Nachname","E-Mail","Telefon","Ort","Dienstleistung"][i]}<input name={x} required /></label>)}</div><button className="primary full">Kunde erstellen</button></form></div> }

createRoot(document.getElementById("root")).render(<App />);