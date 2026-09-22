/* Add-on peserta PASANGAN/KELOMPOK — dimuat SETELAH script utama spectra-admin.html */
(function(){
  var st=document.createElement('style');
  st.textContent='.angRow{display:grid;grid-template-columns:2fr 1fr 1fr auto;gap:.4rem;margin-bottom:.4rem;align-items:center}tr.grpHead td{background:var(--surface-2);font-weight:700}';
  document.head.appendChild(st);

  var ANGK={};
  function angOf(id){return (ANGK[id]&&ANGK[id].anggota)||[];}
  function jenisOf(id){return (ANGK[id]&&ANGK[id].jenis_peserta)||'INDIVIDU';}
  async function loadAnggota(cab,kat,pid){
    var r=await (await sb()).rpc('spectra_list_anggota',{p_kode_akses:sesi.kode_akses,p_cabang_id:cab||null,p_kategori_id:kat||null,p_peserta_id:pid||null});
    if(r.error)return; (r.data||[]).forEach(function(a){ANGK[a.peserta_id]=a;});
  }
  window.angOf=angOf; window.jenisOf=jenisOf; window.loadAnggota=loadAnggota;

  function injectHeader(){
    var tb=document.getElementById('tbPeserta'); if(!tb)return;
    var tr=tb.closest('table').querySelector('thead tr'); if(!tr||tr.querySelector('th[data-jenis]'))return;
    var th=document.createElement('th'); th.textContent='Jenis'; th.setAttribute('data-jenis','1');
    tr.insertBefore(th,tr.lastElementChild);
  }

  var _renderSemua=window.renderSemua;
  window.renderSemua=function(){ loadAnggota(null,null,null).then(function(){ injectHeader(); if(_renderSemua)_renderSemua(); }); };

  window.renderPeserta=function(){
    var q=(el('qPeserta').value||'').toLowerCase(); var fc=el('fPCabang').value;
    var rows=(DATA.peserta||[]).filter(function(p){return (!fc||p.cabang_id===fc)&&(!q||(p.nama_peserta+' '+(p.asal_instansi||'')).toLowerCase().indexOf(q)>=0);});
    el('tbPeserta').innerHTML=rows.map(function(p){
      var j=jenisOf(p.id); var lbl=(j!=='INDIVIDU')?(j+' ('+angOf(p.id).length+')'):'Individu';
      return '<tr><td>'+p.nomor_undian+'</td><td><b>'+esc(p.nama_peserta)+'</b></td><td>'+(p.asal_instansi||'-')+'</td>'+
        '<td class="small muted">'+cabangById(p.cabang_id).nama_cabang+'</td><td class="small muted">'+(kategoriById(p.kategori_id).nama_kategori||'Umum')+'</td>'+
        '<td class="small muted">'+lbl+'</td><td class="small muted">'+((p.tempat_lahir||'')+(p.tanggal_lahir?', '+p.tanggal_lahir:''))+'</td>'+
        '<td class="small muted">'+(p.nisn||'-')+'</td><td class="small muted">'+(p.kelas||'-')+'</td>'+
        '<td><div class="actCell"><button class="btn btn-ghost btn-icon" onclick="openEditPeserta(\''+p.id+'\')"><i data-lucide="pencil"></i></button>'+
        '<button class="btn btn-danger btn-icon" onclick="hapusPeserta(\''+p.id+'\')"><i data-lucide="trash-2"></i></button></div></td></tr>';
    }).join(''); icons();
  };

  window.unduhTemplate=function(){
    var ws=XLSX.utils.aoa_to_sheet([
      ['Nomor Undian','Nama Peserta','Asal Instansi','Cabang','Kategori','Tempat Lahir','Tanggal Lahir','NISN','Kelas','Jenis','Daftar Anggota','Daftar NISN','Daftar Kelas'],
      ['1','Contoh Nama','SD Contoh 1',(DATA.cabang[0]||{}).nama_cabang||'Cabang','Putra','Sragen','2014-05-12','12345678','5A','INDIVIDU','','',''],
      ['2','Tim Contoh','SD Contoh 2',(DATA.cabang[0]||{}).nama_cabang||'Cabang','Putra','','','','','KELOMPOK','Nama1; Nama2; Nama3','N1; N2; N3','5A; 5A; 5B']]);
    XLSX.writeFile({SheetNames:['Peserta'],Sheets:{Peserta:ws}},'template_peserta_spectra.xlsx');
  };

  window.parseFile=function(file){
    if(!file)return; var rd=new FileReader();
    rd.onload=function(e){ try{
      var wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
      var aoa=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''});
      var hi=-1; for(var i=0;i<aoa.length;i++){var hn=false,hm=false;for(var c=0;c<aoa[i].length;c++){if(/nomor/i.test(aoa[i][c]))hn=true;if(/nama/i.test(aoa[i][c]))hm=true;}if(hn&&hm){hi=i;break;}}
      if(hi===-1)return toast('Header tidak ditemukan.','err');
      var heads=aoa[hi].map(function(h){return String(h).toLowerCase();});
      function col(t){for(var i=0;i<heads.length;i++){if(heads[i].indexOf(t)>=0)return i;}return -1;}
      var iN=col('nomor'),iNm=col('nama'); var iA=col('asal'); if(iA<0)iA=col('sekolah'); if(iA<0)iA=col('instansi');
      var iC=col('cabang'),iK=col('kategori'),iT=col('tempat'),iTL=col('tanggal'),iNISN=col('nisn'),iKLS=col('kelas');
      var iJ=col('jenis'),iAng=col('daftar anggota'),iAngN=col('daftar nisn'),iAngK=col('daftar kelas');
      function split(v){return String(v==null?'':v).split(';').map(function(s){return s.trim();}).filter(Boolean);}
      PREV=[]; aoa.slice(hi+1).forEach(function(r,idx){ var empty=true; for(var c=0;c<r.length;c++){if(String(r[c]).trim()!==''){empty=false;break;}} if(empty)return;
        var nomor=String(r[iN]==null?'':r[iN]).trim(), nama=String(r[iNm]==null?'':r[iNm]).trim();
        var cabRef=iC>=0?String(r[iC]==null?'':r[iC]).trim():'';
        var cab=(DATA.cabang||[]).find(function(c){return c.kode_cabang.toLowerCase()===cabRef.toLowerCase()||c.nama_cabang.toLowerCase()===cabRef.toLowerCase();});
        var ok=true,pesan='✅ siap'; if(!nomor||!nama){ok=false;pesan='🔴 nomor/nama kosong';} else if(!cab){ok=false;pesan='🔴 cabang tidak dikenal';}
        var namaA=iAng>=0?split(r[iAng]):[]; var nisnA=iAngN>=0?split(r[iAngN]):[]; var klsA=iAngK>=0?split(r[iAngK]):[];
        var anggota=namaA.map(function(nm,i){return {nama:nm,nisn:nisnA[i]||'',kelas:klsA[i]||''};});
        var jenis=iJ>=0?String(r[iJ]==null?'':r[iJ]).trim().toUpperCase():'INDIVIDU';
        PREV.push({baris:hi+2+idx,nomor:nomor,nama:nama,asal:iA>=0?String(r[iA]==null?'':r[iA]).trim():'',cabang:cabRef,kategori:iK>=0?String(r[iK]==null?'':r[iK]).trim():'',tempat:iT>=0?String(r[iT]==null?'':r[iT]).trim():'',tanggal:iTL>=0?fmtTglExcel(r[iTL]):'',nisn:iNISN>=0?String(r[iNISN]==null?'':r[iNISN]).trim():'',kelas:iKLS>=0?String(r[iKLS]==null?'':r[iKLS]).trim():'',jenis_peserta:jenis,anggota:anggota,ok:ok,pesan:pesan});
      });
      var okN=PREV.filter(function(r){return r.ok;}).length;
      el('imporPrev').hidden=false; el('pvOk').textContent=okN+' valid'; el('pvErr').textContent=(PREV.length-okN)+' error'; el('btnImpor').disabled=(okN===0);
      el('tbPrev').innerHTML=PREV.map(function(r){return '<tr><td>'+r.baris+'</td><td>'+r.nomor+'</td><td>'+r.nama+'</td><td>'+r.cabang+'</td><td>'+(r.kategori||'-')+'</td><td class="small" style="color:'+(r.ok?'var(--ok-text)':'var(--danger-text)')+'">'+r.pesan+(r.jenis_peserta!=='INDIVIDU'?' • '+r.jenis_peserta+' ('+r.anggota.length+')':'')+'</td></tr>';}).join(''); icons();
    }catch(err){ toast('Gagal membaca: '+err.message,'err'); } };
    rd.readAsArrayBuffer(file);
  };

  window.jalankanImport=async function(){
    var rows=PREV.filter(function(r){return r.ok;}).map(function(r){return {nomor_undian:r.nomor,nama_peserta:r.nama,asal_instansi:r.asal,cabang:r.cabang,kategori:r.kategori,tempat_lahir:r.tempat,tanggal_lahir:r.tanggal,nisn:r.nisn,kelas:r.kelas,jenis_peserta:r.jenis_peserta,anggota:r.anggota};});
    var r=await (await sb()).rpc('spectra_import_peserta',{p_kode_akses:sesi.kode_akses,p_rows:rows});
    if(r.error)return toast(r.error.message,'err'); if(r.data.error)return toast(r.data.error,'err');
    toast('Import: '+r.data.inserted+' baru, '+r.data.updated+' diperbarui.','ok'); el('imporPrev').hidden=true; await muat();
  };

  function angRowHtml(a){return '<div class="angRow"><input class="input ang-nama" placeholder="Nama anggota" value="'+esc(a&&a.nama||'')+'"><input class="input ang-nisn" placeholder="NISN" value="'+esc(a&&a.nisn||'')+'"><input class="input ang-kelas" placeholder="Kelas" value="'+esc(a&&a.kelas||'')+'"><button class="btn btn-danger btn-icon" onclick="this.parentElement.remove()"><i data-lucide="x"></i></button></div>';}
  window.addAngRow=function(){ el('angList').insertAdjacentHTML('beforeend',angRowHtml(null)); icons(); };
  function collectAnggota(){ var out=[]; var kids=el('angList').children; for(var i=0;i<kids.length;i++){var d=kids[i]; var nm=d.querySelector('.ang-nama').value.trim(); if(nm)out.push({nama:nm,nisn:d.querySelector('.ang-nisn').value.trim(),kelas:d.querySelector('.ang-kelas').value.trim()});} return out; }

  window.openEditPeserta=async function(id){
    var p=pesertaById(id); if(!p)return; await loadAnggota(null,null,id); var ang=angOf(id), jns=jenisOf(id);
    var ovl=document.createElement('div'); ovl.className='ovl';
    ovl.innerHTML='<div class="card"><h3>Edit Peserta</h3><div class="grid2">'
      +'<label class="field"><span>Nomor Undian</span><input class="input" id="eNomor" value="'+esc(p.nomor_undian)+'"></label>'
      +'<label class="field"><span>Nama</span><input class="input" id="eNama" value="'+esc(p.nama_peserta)+'"></label>'
      +'<label class="field"><span>Asal</span><input class="input" id="eAsal" value="'+esc(p.asal_instansi||'')+'"></label>'
      +'<label class="field"><span>Jenis</span><select class="input" id="eJenis">'+['INDIVIDU','PASANGAN','KELOMPOK'].map(function(j){return '<option '+(j===jns?'selected':'')+'>'+j+'</option>';}).join('')+'</select></label>'
      +'<label class="field"><span>Cabang</span><select class="input" id="eCabang">'+DATA.cabang.map(function(c){return '<option value="'+c.id+'" '+(p.cabang_id===c.id?'selected':'')+'>'+esc(c.nama_cabang)+'</option>';}).join('')+'</select></label>'
      +'<label class="field"><span>Kategori</span><select class="input" id="eKategori"><option value="">(Tanpa)</option>'+DATA.kategori.map(function(k){return '<option value="'+k.id+'" '+(p.kategori_id===k.id?'selected':'')+'>'+esc(k.nama_kategori)+'</option>';}).join('')+'</select></label>'
      +'<label class="field"><span>Tempat Lahir</span><input class="input" id="eTempat" value="'+esc(p.tempat_lahir||'')+'"></label>'
      +'<label class="field"><span>Tanggal Lahir</span><input class="input" type="date" id="eTgl" value="'+(p.tanggal_lahir||'')+'"></label>'
      +'<label class="field"><span>NISN</span><input class="input" id="eNisn" value="'+esc(p.nisn||'')+'"></label>'
      +'<label class="field"><span>Kelas</span><input class="input" id="eKelas" value="'+esc(p.kelas||'')+'"></label></div>'
      +'<label class="field"><span>Anggota (PASANGAN/KELOMPOK)</span></label><div id="angList">'+ang.map(function(a){return angRowHtml(a);}).join('')+'</div>'
      +'<button class="btn btn-ghost" onclick="addAngRow()">+ Tambah Anggota</button>'
      +'<div class="modal-actions"><button class="btn btn-ghost" onclick="this.closest(\'.ovl\').remove()">Batal</button><button class="btn btn-primary" onclick="simpanEditPeserta(\''+p.id+'\')">Simpan</button></div></div>';
    document.body.appendChild(ovl); icons();
  };

  window.simpanEditPeserta=async function(id){
    var payload={nomor_undian:el('eNomor').value.trim(),nama_peserta:el('eNama').value.trim(),asal_instansi:el('eAsal').value.trim(),jenis_peserta:el('eJenis').value,cabang_id:el('eCabang').value,kategori_id:el('eKategori').value||'',tempat_lahir:el('eTempat').value.trim(),tanggal_lahir:el('eTgl').value,nisn:el('eNisn').value.trim(),kelas:el('eKelas').value.trim(),anggota:collectAnggota()};
    var r=await (await sb()).rpc('spectra_update_peserta',{p_kode_akses:sesi.kode_akses,p_peserta_id:id,p_payload:payload});
    if(r.error)return toast(r.error.message,'err'); if(r.data.error)return toast(r.data.error,'err');
    document.querySelector('.ovl').remove(); toast('Peserta diperbarui.','ok'); await muat();
  };

  window.cetakDaftarHadir=function(){
    var fc=el('fDCabang').value; if(!fc)return toast('Pilih cabang.','err');
    var fk=el('fDKategori').value; var rows=daftarRows(); var cab=cabangById(fc); var orient=el('fDOrient').value||'portrait';
    var op=(DATA.petugas||[]).find(function(p){return p.peran==='OPERATOR'&&p.cabang_id===fc&&(!fk||p.kategori_id===fk);})||(DATA.petugas||[]).find(function(p){return p.peran==='OPERATOR'&&p.cabang_id===fc;});
    var h=kopHtml()+'<div class="judul"><div class="t1">Daftar Hadir Peserta</div><div class="t2">'+esc(DATA.event.nama)+' — Cabang '+esc(cab.nama_cabang)+(fk?(' — '+esc(kategoriById(fk).nama_kategori)):'')+'</div></div>';
    h+='<table class="data"><thead><tr><th style="width:34px">No</th><th style="width:56px">Undian</th><th>Nama Peserta</th><th>Asal Sekolah</th><th style="width:46px">Kelas</th><th style="width:200px">Tanda Tangan</th></tr></thead><tbody>';
    var no=0; rows.forEach(function(r){ var j=jenisOf(r.p.id), ang=angOf(r.p.id);
      if(j!=='INDIVIDU'&&ang.length){ no++; h+='<tr class="grpHead"><td style="text-align:center">'+no+'</td><td style="text-align:center">'+esc(r.p.nomor_undian)+'</td><td colspan="4">'+esc(r.p.nama_peserta)+' ('+j+' — '+ang.length+' anggota)</td></tr>';
        ang.forEach(function(a){ h+='<tr><td></td><td></td><td>'+esc(a.nama)+'</td><td>'+esc(r.p.asal_instansi||'-')+'</td><td style="text-align:center">'+esc(a.kelas||'')+'</td><td style="height:32px"></td></tr>'; }); }
      else { no++; h+='<tr><td style="text-align:center">'+no+'</td><td style="text-align:center">'+esc(r.p.nomor_undian)+'</td><td>'+esc(r.p.nama_peserta)+'</td><td>'+esc(r.p.asal_instansi||'-')+'</td><td style="text-align:center">'+esc(r.p.kelas||'-')+'</td><td style="height:36px"></td></tr>'; } });
    h+='</tbody></table>';
    var d=new Date(); var B=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
    h+='<div style="margin-top:26px;display:flex;justify-content:space-between;gap:40px"><div style="text-align:center;flex:1">Mengetahui,<br>Ketua Panitia,<br><br><br><br><b>'+esc(SET().ketua_panitia||'(..................)')+'</b></div><div style="text-align:center;flex:1">'+esc(SET().tempat_cetak||'')+', '+d.getDate()+' '+B[d.getMonth()]+' '+d.getFullYear()+'<br>Operator Cabang,<br><br><br><br><b>'+esc(op?op.nama_petugas:'(..................)')+'</b></div></div>';
    var css=PRINT_CSS; if(orient==='landscape')css=css.replace('@page{size:A4;','@page{size:A4 landscape;');
    var w=window.open('','_blank'); w.document.write('<html><head><meta charset="utf-8">'+css+'</head><body>'+h+'</body></html>'); w.document.close(); w.focus(); setTimeout(function(){w.print();},300);
  };
})();