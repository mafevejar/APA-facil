import { useState, useRef, useEffect } from "react";
import "./App.css";

import html2pdf from "html2pdf.js";
import mammoth from "mammoth";

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  ImageRun,
  AlignmentType,
} from "docx";

function App() {
  const [vista, setVista] = useState("inicio");

  const [zoom, setZoom] = useState(100);

  const [mostrarRegla, setMostrarRegla] = useState(true);

  const [numeracionActiva, setNumeracionActiva] = useState(true);

  const [paginas, setPaginas] = useState([1]);

  const [tamanoActual, setTamanoActual] = useState("12");

  const [interlineadoActual, setInterlineadoActual] = useState("2");

  const editorRefs = useRef([]);

  const ultimoParrafo = useRef(null);

  const seleccionGuardada = useRef(null);

  /* =========================================================
     GUARDAR SELECCIÓN DEL EDITOR
     ========================================================= */

  const guardarSeleccion = () => {
    const seleccion = window.getSelection();

    if (!seleccion || seleccion.rangeCount === 0) {
      return;
    }

    const range = seleccion.getRangeAt(0);

    const editor = editorRefs.current.find(
      (elemento) =>
        elemento && elemento.contains(range.commonAncestorContainer),
    );

    if (!editor) return;

    seleccionGuardada.current = range.cloneRange();
  };

  /* =========================================================
     RESTAURAR SELECCIÓN
     ========================================================= */

  const restaurarSeleccion = () => {
    const range = seleccionGuardada.current;

    if (!range) return;

    const seleccion = window.getSelection();

    seleccion.removeAllRanges();

    try {
      seleccion.addRange(range);
    } catch (error) {
      console.warn("No se pudo restaurar la selección.");
    }
  };

  /* =========================================================
     FORMATO GENERAL
     ========================================================= */

  const aplicarFormato = (comando, valor = null) => {
    const editor = editorRefs.current[paginas.length - 1];

    if (!editor) return;

    editor.focus();

    restaurarSeleccion();

    try {
      document.execCommand("styleWithCSS", false, true);
    } catch (error) {}

    document.execCommand(comando, false, valor);

    editor.focus();

    guardarSeleccion();
  };

  /* =========================================================
     TAMAÑO DE LETRA
     ========================================================= */

  const cambiarTamanoLetra = (tamano) => {
    const editor = editorRefs.current[paginas.length - 1];

    if (!editor) return;

    editor.focus();

    restaurarSeleccion();

    const seleccion = window.getSelection();

    if (!seleccion || seleccion.rangeCount === 0) {
      return;
    }

    const range = seleccion.getRangeAt(0);

    /*
     * TEXTO SELECCIONADO
     */

    if (!seleccion.isCollapsed) {
      const span = document.createElement("span");

      span.style.fontSize = `${tamano}pt`;

      try {
        range.surroundContents(span);
      } catch (error) {
        const contenido = range.extractContents();

        span.appendChild(contenido);

        range.insertNode(span);
      }

      const nuevoRango = document.createRange();

      nuevoRango.selectNodeContents(span);

      seleccion.removeAllRanges();

      seleccion.addRange(nuevoRango);

      seleccionGuardada.current = nuevoRango.cloneRange();
    } else {
      /*
       * SIN SELECCIÓN:
       * LO QUE SE ESCRIBA DESPUÉS
       * TENDRÁ EL TAMAÑO ELEGIDO.
       */

      const span = document.createElement("span");

      span.style.fontSize = `${tamano}pt`;

      const texto = document.createTextNode("\u200B");

      span.appendChild(texto);

      range.insertNode(span);

      const nuevoRango = document.createRange();

      nuevoRango.setStart(texto, 1);

      nuevoRango.collapse(true);

      seleccion.removeAllRanges();

      seleccion.addRange(nuevoRango);

      seleccionGuardada.current = nuevoRango.cloneRange();
    }

    setTamanoActual(String(tamano));

    editor.focus();
  };

  /* =========================================================
     INTERLINEADO
     ========================================================= */

  const cambiarInterlineado = (valor) => {
    setInterlineadoActual(valor);

    editorRefs.current.forEach((editor) => {
      if (!editor) return;

      editor.style.lineHeight = valor;

      editor.querySelectorAll("p, h2, h3, h4, li").forEach((elemento) => {
        elemento.style.lineHeight = valor;
      });
    });
  };

  /* =========================================================
     COLOR DE TEXTO
     ========================================================= */

  const cambiarColorTexto = (color) => {
    aplicarFormato("foreColor", color);
  };

  /* =========================================================
     RESALTADO
     ========================================================= */

  const cambiarResaltado = (color) => {
    const editor = editorRefs.current[paginas.length - 1];

    if (!editor) return;

    editor.focus();

    restaurarSeleccion();

    try {
      document.execCommand("styleWithCSS", false, true);
    } catch (error) {}

    document.execCommand("hiliteColor", false, color);

    editor.focus();

    guardarSeleccion();
  };

  /* =========================================================
     GUARDAR DOCUMENTO
     ========================================================= */

  const guardarDocumento = () => {
    const documentos = editorRefs.current.map((editor) =>
      editor ? editor.innerHTML : "",
    );

    localStorage.setItem("documentoAPA", JSON.stringify(documentos));

    alert("Documento guardado correctamente 💜");
  };

  /* =========================================================
     CARGAR DOCUMENTO GUARDADO
     ========================================================= */

  useEffect(() => {
    if (vista !== "editor") {
      return;
    }

    const documentoGuardado = localStorage.getItem("documentoAPA");

    if (!documentoGuardado) {
      return;
    }

    try {
      const documentos = JSON.parse(documentoGuardado);

      if (Array.isArray(documentos)) {
        setPaginas(
          documentos.length > 0 ? documentos.map((_, index) => index + 1) : [1],
        );

        setTimeout(() => {
          documentos.forEach((contenido, index) => {
            if (editorRefs.current[index]) {
              editorRefs.current[index].innerHTML = contenido;
            }
          });
        }, 150);
      }
    } catch (error) {
      console.error("No se pudo cargar el documento:", error);
    }
  }, [vista]);

  /* =========================================================
     EXPORTAR PDF
     ========================================================= */

  const exportarPDF = async () => {
    const paginasPDF = editorRefs.current.filter((editor) => editor);

    if (!paginasPDF.length) {
      return;
    }

    const contenedor = document.createElement("div");

    contenedor.style.background = "white";

    paginasPDF.forEach((editor, index) => {
      const pagina = document.createElement("div");

      pagina.style.width = "21cm";

      pagina.style.minHeight = "29.7cm";

      pagina.style.padding = "2.54cm";

      pagina.style.boxSizing = "border-box";

      pagina.style.background = "white";

      pagina.style.color = "#000";

      pagina.style.fontFamily = "'Times New Roman', serif";

      pagina.style.fontSize = "12pt";

      pagina.style.lineHeight = "2";

      pagina.style.position = "relative";

      pagina.style.pageBreakAfter =
        index < paginasPDF.length - 1 ? "always" : "auto";

      if (numeracionActiva) {
        const numero = document.createElement("div");

        numero.textContent = index + 1;

        numero.style.position = "absolute";

        numero.style.top = "1cm";

        numero.style.right = "1.27cm";

        numero.style.fontFamily = "'Times New Roman', serif";

        numero.style.fontSize = "12pt";

        numero.style.color = "#000";

        pagina.appendChild(numero);
      }

      contenido.querySelectorAll("h2, h3, h4").forEach((h) => {
        h.style.fontSize = "12pt";
        h.style.margin = "0";
        h.style.textIndent = "0";
        h.style.fontWeight = "bold";
        h.style.fontStyle = h.tagName === "H4" ? "italic" : "normal";
        h.style.textAlign = h.tagName === "H2" ? "center" : "left";
      });
    });

    document.body.appendChild(contenedor);

    const opciones = {
      margin: 0,

      filename: "Documento_APA.pdf",

      image: {
        type: "jpeg",
        quality: 0.98,
      },

      html2canvas: {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
      },

      jsPDF: {
        unit: "cm",
        format: "a4",
        orientation: "portrait",
      },

      pagebreak: {
        mode: ["css", "legacy"],
      },
    };

    try {
      await html2pdf().set(opciones).from(contenedor).save();
    } catch (error) {
      console.error("Error exportando PDF:", error);

      alert("No se pudo exportar el PDF.");
    }

    document.body.removeChild(contenedor);
  };

  /* =========================================================
     CONVERTIR HTML A WORD
     ========================================================= */

  const convertirTextoWord = (elemento, estilosBase = {}) => {
    const runs = [];

    const recorrer = (nodo, estilos = estilosBase) => {
      if (nodo.nodeType === Node.TEXT_NODE) {
        if (nodo.textContent) {
          runs.push(
            new TextRun({
              text: nodo.textContent,
              bold: estilos.bold,
              italics: estilos.italics,
              color: estilos.color,
              font: "Times New Roman",
              size: 24, // 12pt
            }),
          );
        }
        return;
      }

      if (nodo.nodeType !== Node.ELEMENT_NODE) {
        return;
      }

      let color = estilos.color;

      if (nodo.style && nodo.style.color) {
        color = nodo.style.color.replace("#", "").toUpperCase();
      }

      const nuevosEstilos = {
        bold: estilos.bold || nodo.tagName === "STRONG" || nodo.tagName === "B",
        italics:
          estilos.italics || nodo.tagName === "EM" || nodo.tagName === "I",
        color,
      };

      nodo.childNodes.forEach((hijo) => recorrer(hijo, nuevosEstilos));
    };

    elemento.childNodes.forEach((hijo) => recorrer(hijo, estilosBase));

    return runs;
  };

  /* =========================================================
     CREAR CONTENIDO WORD
     ========================================================= */

  const crearContenidoWord = async (editor) => {
    const elementos = Array.from(editor.children);

    const contenido = [];

    for (const elemento of elementos) {
      const tag = elemento.tagName.toLowerCase();

      if (tag === "h2") {
        contenido.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 480, after: 0 },
            children: convertirTextoWord(elemento, { bold: true }),
          }),
        );
        continue;
      }

      if (tag === "h3") {
        contenido.push(
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { line: 480, after: 0 },
            children: convertirTextoWord(elemento, { bold: true }),
          }),
        );
        continue;
      }

      if (tag === "h4") {
        contenido.push(
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { line: 480, after: 0 },
            children: convertirTextoWord(elemento, {
              bold: true,
              italics: true,
            }),
          }),
        );
        continue;
      }

      if (tag === "p") {
        contenido.push(
          new Paragraph({
            alignment: AlignmentType.LEFT,

            indent: {
              firstLine:
                elemento.style.textIndent &&
                elemento.style.textIndent !== "0px" &&
                elemento.style.textIndent !== "0"
                  ? 720
                  : 0,
            },

            spacing: {
              line: 480,
              after: 0,
            },

            children: convertirTextoWord(elemento),
          }),
        );

        continue;
      }

      if (tag === "img") {
        try {
          const src = elemento.getAttribute("src");

          if (src && src.startsWith("data:image")) {
            const base64 = src.split(",")[1];

            const binary = atob(base64);

            const bytes = new Uint8Array(binary.length);

            for (let i = 0; i < binary.length; i++) {
              bytes[i] = binary.charCodeAt(i);
            }

            contenido.push(
              new Paragraph({
                alignment: AlignmentType.CENTER,

                children: [
                  new ImageRun({
                    data: bytes,

                    transformation: {
                      width: 400,

                      height: 300,
                    },
                  }),
                ],
              }),
            );
          }
        } catch (error) {
          console.error("No se pudo convertir imagen:", error);
        }

        continue;
      }

      if (tag === "table") {
        const filas = Array.from(elemento.querySelectorAll("tr"));

        const filasWord = filas.map((fila) => {
          const celdas = Array.from(fila.querySelectorAll("th, td"));

          return new TableRow({
            children: celdas.map(
              (celda) =>
                new TableCell({
                  children: [
                    new Paragraph({
                      children: convertirTextoWord(celda),
                    }),
                  ],
                }),
            ),
          });
        });

        contenido.push(
          new Table({
            rows: filasWord,

            width: {
              size: 100,
              type: "pct",
            },
          }),
        );

        continue;
      }

      if (elemento.textContent.trim()) {
        contenido.push(
          new Paragraph({
            spacing: {
              line: 480,
            },

            children: convertirTextoWord(elemento),
          }),
        );
      }
    }

    return contenido;
  };

  /* =========================================================
     EXPORTAR WORD
     ========================================================= */

  const exportarWord = async () => {
    const editores = editorRefs.current.filter((editor) => editor);

    if (!editores.length) {
      return;
    }

    try {
      const secciones = [];

      for (let index = 0; index < editores.length; index++) {
        const contenido = await crearContenidoWord(editores[index]);

        if (numeracionActiva) {
          contenido.unshift(
            new Paragraph({
              alignment: AlignmentType.RIGHT,

              spacing: {
                after: 0,
              },

              children: [
                new TextRun({
                  text: String(index + 1),

                  font: "Times New Roman",

                  size: 24,
                }),
              ],
            }),
          );
        }

        secciones.push({
          properties: {
            page: {
              size: {
                width: 11906,

                height: 16838,
              },

              margin: {
                top: 1440,

                right: 1440,

                bottom: 1440,

                left: 1440,
              },
            },
          },

          children: contenido,
        });
      }

      const documento = new Document({
        sections: secciones,
      });

      const blob = await Packer.toBlob(documento);

      const url = URL.createObjectURL(blob);

      const enlace = document.createElement("a");

      enlace.href = url;

      enlace.download = "Documento_APA.docx";

      document.body.appendChild(enlace);

      enlace.click();

      document.body.removeChild(enlace);

      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Error exportando DOCX:", error);

      alert("No se pudo exportar el documento Word.");
    }
  };

  /* =========================================================
     CARGAR DOCUMENTO
     ========================================================= */

  const cargarDocumento = async (e) => {
    const archivo = e.target.files[0];

    if (!archivo) {
      return;
    }

    try {
      const extension = archivo.name.split(".").pop().toLowerCase();

      /* TXT */

      if (extension === "txt") {
        const texto = await archivo.text();

        const html = texto
          .split(/\r?\n/)
          .map((linea) => (linea.trim() ? `<p>${linea}</p>` : `<p><br></p>`))
          .join("");

        localStorage.setItem("documentoAPA", JSON.stringify([html]));

        setPaginas([1]);

        setVista("editor");

        e.target.value = "";

        return;
      }

      /* DOCX */

      if (extension === "docx") {
        const arrayBuffer = await archivo.arrayBuffer();

        const resultado = await mammoth.convertToHtml({
          arrayBuffer,
        });

        const html = resultado.value;

        localStorage.setItem("documentoAPA", JSON.stringify([html]));

        setPaginas([1]);

        setVista("editor");

        e.target.value = "";

        return;
      }

      alert("Puedes cargar documentos .docx o .txt.");
    } catch (error) {
      console.error("Error cargando documento:", error);

      alert("No se pudo cargar el documento.");
    }

    e.target.value = "";
  };

  /* =========================================================
     NUEVA PÁGINA
     ========================================================= */

  const agregarPagina = () => {
    setPaginas((actuales) => [...actuales, actuales.length + 1]);
  };

  /* =========================================================
     ELIMINAR PÁGINA
     ========================================================= */

  const eliminarPagina = () => {
    if (paginas.length <= 1) {
      alert("El documento debe tener al menos una página.");

      return;
    }

    setPaginas((actuales) => actuales.slice(0, -1));
  };

  /* =========================================================
     ZOOM
     ========================================================= */

  const aumentarZoom = () => {
    setZoom((valor) => Math.min(valor + 10, 160));
  };

  const disminuirZoom = () => {
    setZoom((valor) => Math.max(valor - 10, 60));
  };

  /* =========================================================
     INSERTAR TABLA
     ========================================================= */

  const insertarTabla = (filas, columnas) => {
    const editor = editorRefs.current[paginas.length - 1];

    if (!editor) return;

    editor.focus();

    restaurarSeleccion();

    let tabla = `
      <table class="apa-table">
        <tbody>
    `;

    for (let i = 0; i < filas; i++) {
      tabla += "<tr>";

      for (let j = 0; j < columnas; j++) {
        tabla += `
          <td contenteditable="true">
            Escriba aquí
          </td>
        `;
      }

      tabla += "</tr>";
    }

    tabla += `
        </tbody>
      </table>
      <p><br></p>
    `;

    document.execCommand("insertHTML", false, tabla);
  };

  /* =========================================================
     SELECCIONAR TABLA
     ========================================================= */

  const seleccionarTabla = () => {
    const filas = prompt("¿Cuántas filas desea?");

    const columnas = prompt("¿Cuántas columnas desea?");

    if (!filas || !columnas) {
      return;
    }

    const numeroFilas = parseInt(filas);

    const numeroColumnas = parseInt(columnas);

    if (
      numeroFilas < 1 ||
      numeroColumnas < 1 ||
      numeroFilas > 15 ||
      numeroColumnas > 10
    ) {
      alert("Seleccione entre 1 y 15 filas y entre 1 y 10 columnas.");

      return;
    }

    insertarTabla(numeroFilas, numeroColumnas);
  };

  /* =========================================================
     INSERTAR IMAGEN
     ========================================================= */

  const insertarImagen = (e) => {
    const archivo = e.target.files[0];

    if (!archivo) return;

    if (!archivo.type.startsWith("image/")) {
      alert("Seleccione una imagen válida.");

      return;
    }

    const lector = new FileReader();

    lector.onload = (evento) => {
      const editor = editorRefs.current[paginas.length - 1];

      if (!editor) return;

      editor.focus();

      restaurarSeleccion();

      const imagen = `

            <img
              src="${evento.target.result}"
              class="apa-image"
              style="
                width: 300px;
                max-width: 100%;
                height: auto;
              "
            />

            <p><br></p>

          `;

      document.execCommand("insertHTML", false, imagen);
    };

    lector.readAsDataURL(archivo);

    e.target.value = "";
  };

  /* =========================================================
     ACTUALIZAR ÚLTIMO PÁRRAFO
     ========================================================= */

  const actualizarParrafo = (e) => {
    const parrafo = e.target.closest("p");

    if (parrafo) {
      ultimoParrafo.current = parrafo;
    }

    guardarSeleccion();
  };

  /* =========================================================
     SANGRÍA
     ========================================================= */

  const aplicarSangria = () => {
    const editor = editorRefs.current[paginas.length - 1];
    if (!editor) return;

    editor.focus();
    restaurarSeleccion();

    const seleccion = window.getSelection();
    if (!seleccion || seleccion.rangeCount === 0) return;

    const range = seleccion.getRangeAt(0);

    // Encuentra el bloque (p, li, h2, h3, h4) de un nodo
    const obtenerBloque = (nodo) => {
      if (nodo && nodo.nodeType === Node.TEXT_NODE) {
        nodo = nodo.parentElement;
      }
      return nodo?.closest("p, li, h2, h3, h4");
    };

    // Todos los bloques tocados por la selección
    const bloques = new Set();
    const inicio = obtenerBloque(range.startContainer);
    const fin = obtenerBloque(range.endContainer);

    if (inicio) bloques.add(inicio);
    if (fin) bloques.add(fin);

    editor.querySelectorAll("p, li, h2, h3, h4").forEach((b) => {
      if (range.intersectsNode(b)) bloques.add(b);
    });

    if (bloques.size === 0) return;

    // Si el primer bloque ya tiene sangría, se quita en todos; si no, se pone
    const primero = Array.from(bloques)[0];
    const yaTieneSangria =
      primero.style.textIndent &&
      primero.style.textIndent !== "0px" &&
      primero.style.textIndent !== "0";

    bloques.forEach((b) => {
      b.style.textIndent = yaTieneSangria ? "0" : "1.27cm";
      b.style.marginLeft = "0";
    });

    guardarSeleccion();
  };

  /* =========================================================
     PEGAR TEXTO LIMPIO
     ========================================================= */

  const pegarTexto = (e) => {
    e.preventDefault();

    const texto = e.clipboardData.getData("text/plain");

    if (!texto) return;

    const editor = editorRefs.current[paginas.length - 1];

    if (!editor) return;

    editor.focus();

    restaurarSeleccion();

    const lineas = texto.split(/\r?\n/);

    const html = lineas
      .map((linea) => {
        const textoSeguro = linea
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");

        return `
                <p
                  style="
                    text-indent: 1.27cm;
                    margin-left: 0;
                  "
                >
                  ${textoSeguro}
                </p>
              `;
      })
      .join("");

    document.execCommand("insertHTML", false, html);

    guardarSeleccion();
  };

  /* =========================================================
     TECLADO DEL EDITOR
     ========================================================= */

  const manejarKeyDown = (e) => {
    guardarSeleccion();

    /*
     * TAB = sangría
     */

    if (e.key === "Tab") {
      e.preventDefault();
      aplicarSangria();
    }
  };

  /* =========================================================
     EDITOR
     ========================================================= */

  if (vista === "editor") {
    return (
      <div className="app">
        {/* NAVBAR */}

        <header className="navbar">
          <div className="logo" onClick={() => setVista("inicio")}>
            APA Fácil
          </div>

          <nav>
            <button onClick={() => setVista("inicio")}>🏠 Inicio</button>

            <button onClick={guardarDocumento}>💾 Guardar</button>

            <button onClick={exportarWord}>📄 Word</button>

            <button onClick={exportarPDF}>📕 PDF</button>
          </nav>
        </header>

        {/* EDITOR */}

        <main className="editor-page">
          <div className="editor-header">
            <div>
              <h1>Nuevo documento</h1>

              <p>Configura y edita tu documento académico con normas APA.</p>
            </div>

            <button
              className="btn-secondary"
              onClick={() => setVista("inicio")}
            >
              ← Volver
            </button>
          </div>

          <div className="editor-layout">
            {/* PANEL APA */}

            <aside className="apa-panel">
              <h2>⚙️ Configuración APA</h2>

              <label>Márgenes</label>

              <select
                defaultValue="2.54"
                onChange={(e) => {
                  editorRefs.current.forEach((editor) => {
                    if (editor) {
                      editor.style.padding = `${e.target.value}cm`;
                    }
                  });
                }}
              >
                <option value="2.54">2,54 cm</option>

                <option value="2">2 cm</option>

                <option value="3">3 cm</option>
              </select>

              <label>Fuente</label>

              <select
                defaultValue="Times New Roman"
                onChange={(e) => aplicarFormato("fontName", e.target.value)}
              >
                <option value="Times New Roman">Times New Roman</option>

                <option value="Arial">Arial</option>

                <option value="Calibri">Calibri</option>
              </select>

              <label>Tamaño</label>

              <select
                value={tamanoActual}
                onChange={(e) => cambiarTamanoLetra(e.target.value)}
              >
                <option value="14">14</option>

                <option value="12">12</option>

                <option value="11">11</option>

                <option value="10">10</option>
              </select>

              <label>Interlineado</label>

              <select
                value={interlineadoActual}
                onChange={(e) => cambiarInterlineado(e.target.value)}
              >
                <option value="2">Doble</option>

                <option value="1.5">1.5</option>

                <option value="1">Sencillo</option>
              </select>

              <h3>✨ Formato</h3>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("bold");
                }}
              >
                <span>𝐁</span>
                Negrita
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("italic");
                }}
              >
                <span>𝐼</span>
                Cursiva
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("underline");
                }}
              >
                <span>U̲</span>
                Subrayado
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("justifyLeft");
                }}
              >
                <span>≡</span>
                Alinear izquierda
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("justifyCenter");
                }}
              >
                <span>≡</span>
                Centrar
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("justifyRight");
                }}
              >
                <span>≡</span>
                Alinear derecha
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("justifyFull");
                }}
              >
                <span>≡</span>
                Justificar
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarSangria();
                }}
              >
                <span>⇥</span>
                Sangría 1,27 cm
              </button>

              <h3>🎨 Color</h3>

              <label>Color de texto</label>

              <input
                className="color-picker"
                type="color"
                defaultValue="#111827"
                onChange={(e) => cambiarColorTexto(e.target.value)}
              />

              <label>Resaltado</label>

              <input
                className="color-picker"
                type="color"
                defaultValue="#fff59d"
                onChange={(e) => cambiarResaltado(e.target.value)}
              />

              <h3>🔢 Títulos APA</h3>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("formatBlock", "h2");
                }}
              >
                Título 1
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("formatBlock", "h3");
                }}
              >
                Título 2
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("formatBlock", "h4");
                }}
              >
                Título 3
              </button>

              <h3>📋 Listas</h3>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("insertOrderedList");
                }}
              >
                🔢 Lista numerada
              </button>

              <button
                className="format-button"
                onMouseDown={(e) => {
                  e.preventDefault();

                  aplicarFormato("insertUnorderedList");
                }}
              >
                ☷ Lista con viñetas
              </button>

              <h3>🧩 Insertar</h3>

              <button className="format-button" onClick={seleccionarTabla}>
                📊 Insertar tabla
              </button>

              <label className="format-button image-button">
                🖼️ Insertar imagen
                <input
                  type="file"
                  accept="image/*"
                  hidden
                  onChange={insertarImagen}
                />
              </label>
            </aside>

            {/* DOCUMENTO */}

            <section className="document-area">
              {/* TOOLBAR */}

              <div className="toolbar">
                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("undo");
                  }}
                  title="Deshacer"
                >
                  ↶
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("redo");
                  }}
                  title="Rehacer"
                >
                  ↷
                </button>

                <span className="toolbar-separator" />

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("bold");
                  }}
                  title="Negrita"
                >
                  <b>B</b>
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("italic");
                  }}
                  title="Cursiva"
                >
                  <i>I</i>
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("underline");
                  }}
                  title="Subrayado"
                >
                  <u>U</u>
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("justifyLeft");
                  }}
                  title="Alinear izquierda"
                >
                  ≡
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("justifyCenter");
                  }}
                  title="Centrar"
                >
                  ☰
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("justifyRight");
                  }}
                  title="Alinear derecha"
                >
                  ≡
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("justifyFull");
                  }}
                  title="Justificar"
                >
                  ☷
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("insertOrderedList");
                  }}
                  title="Lista numerada"
                >
                  1.
                </button>

                <button
                  onMouseDown={(e) => {
                    e.preventDefault();

                    aplicarFormato("insertUnorderedList");
                  }}
                  title="Lista"
                >
                  •
                </button>

                <span className="toolbar-separator" />

                <button onClick={agregarPagina} title="Nueva página">
                  📄+
                </button>

                <button onClick={eliminarPagina} title="Eliminar página">
                  🗑️
                </button>

                <span className="toolbar-separator" />

                <button
                  className={mostrarRegla ? "toolbar-active" : ""}
                  onClick={() => setMostrarRegla(!mostrarRegla)}
                  title="Mostrar u ocultar regla"
                >
                  📏
                </button>

                <button
                  className={numeracionActiva ? "toolbar-active" : ""}
                  onClick={() => setNumeracionActiva(!numeracionActiva)}
                  title="Numeración de páginas"
                >
                  #
                </button>

                <span className="toolbar-separator" />

                <button onClick={disminuirZoom} title="Alejar">
                  −
                </button>

                <span className="zoom-value">{zoom}%</span>

                <button onClick={aumentarZoom} title="Acercar">
                  +
                </button>
              </div>

              {/* REGLA */}

              {mostrarRegla && (
                <div className="ruler">
                  {Array.from(
                    {
                      length: 17,
                    },
                    (_, index) => (
                      <div className="ruler-number" key={index}>
                        {index}
                      </div>
                    ),
                  )}
                </div>
              )}

              {/* PÁGINAS */}

              <div
                className="pages-container"
                style={{
                  transform: `scale(${zoom / 100})`,
                  transformOrigin: "top center",
                }}
              >
                {paginas.map((numero, index) => (
                  <div className="document-page-wrapper" key={numero}>
                    {numeracionActiva && (
                      <div className="page-number">{numero}</div>
                    )}

                    <div
                      ref={(elemento) => {
                        editorRefs.current[index] = elemento;
                      }}
                      className="document"
                      contentEditable="true"
                      suppressContentEditableWarning={true}
                      data-page-number={numero}
                      onClick={actualizarParrafo}
                      onKeyUp={actualizarParrafo}
                      onMouseUp={guardarSeleccion}
                      onKeyDown={manejarKeyDown}
                      onPaste={pegarTexto}
                    >
                      {index === 0 ? (
                        <>
                          <h2>Escriba el título de su documento</h2>

                          <p>
                            Comience a escribir aquí el contenido de su
                            documento académico.
                          </p>

                          <h3>Introducción</h3>

                          <p>
                            Escriba aquí la introducción de su trabajo. La
                            plataforma permitirá aplicar configuraciones básicas
                            de las normas APA al contenido del documento.
                          </p>
                        </>
                      ) : (
                        <p>Continúe escribiendo aquí...</p>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </section>
          </div>
        </main>
      </div>
    );
  }

  /* =========================================================
     PÁGINA PRINCIPAL
     ========================================================= */

  return (
    <div className="app">
      <header className="navbar">
        <div className="logo">APA Fácil</div>

        <nav>
          <button onClick={() => setVista("inicio")}>🏠 Inicio</button>

          <button onClick={() => setVista("editor")}>✏️ Crear documento</button>

          <label className="nav-upload">
            📂 Cargar documento
            <input
              type="file"
              accept=".docx,.txt"
              onChange={cargarDocumento}
              hidden
            />
          </label>
        </nav>
      </header>

      <main>
        {/* =====================================================
           HERO
           ===================================================== */}

        <section className="hero">
          <div className="hero-badge">✦ Herramienta académica inteligente</div>

          <h1>
            Organiza tus documentos
            <br />
            <span>con normas APA</span>
          </h1>

          <p>
            APA Fácil es una plataforma web diseñada para ayudarte a aplicar
            configuraciones básicas de las normas APA 7.ª edición de manera
            sencilla, rápida y organizada.
          </p>

          <div className="hero-buttons">
            <button className="btn-primary" onClick={() => setVista("editor")}>
              ✨ Crear nuevo documento
            </button>

            <label className="btn-secondary">
              📂 Cargar documento
              <input
                type="file"
                accept=".docx,.txt"
                onChange={cargarDocumento}
                hidden
              />
            </label>
          </div>
        </section>

        {/* =====================================================
           CARACTERÍSTICAS
           ===================================================== */}

        <section className="features">
          <div className="feature-card">
            <span>📄</span>

            <h2>Crear</h2>

            <p>
              Comienza un documento académico desde cero utilizando un editor
              sencillo y organizado.
            </p>
          </div>

          <div className="feature-card">
            <span>📝</span>

            <h2>Formatear</h2>

            <p>
              Aplica configuraciones básicas de las normas APA directamente
              sobre tu documento.
            </p>
          </div>

          <div className="feature-card">
            <span>💾</span>

            <h2>Exportar</h2>

            <p>
              Guarda tu trabajo y expórtalo en formatos PDF o Word para
              continuar trabajando en otras herramientas.
            </p>
          </div>
        </section>

        {/* =====================================================
           COMO USAR
           ===================================================== */}

        <section className="info-section">
          <div className="section-heading">
            <span>🚀</span>

            <p>PASO A PASO</p>

            <h2>¿Cómo usar APA Fácil?</h2>

            <div className="heading-line" />

            <article className="section-intro">
              Sigue estos pasos para crear y organizar tu documento académico.
            </article>
          </div>

          <div className="steps-grid">
            <div className="step-card">
              <div className="step-number">01</div>

              <div className="step-icon">📄</div>

              <h3>Crea o carga</h3>

              <p>
                Crea un documento nuevo desde la página principal o carga un
                archivo Word en formato .docx.
              </p>
            </div>

            <div className="step-card">
              <div className="step-number">02</div>

              <div className="step-icon">⚙️</div>

              <h3>Configura</h3>

              <p>
                Utiliza el panel de configuración para elegir márgenes, fuente,
                tamaño, interlineado y otros elementos de formato.
              </p>
            </div>

            <div className="step-card">
              <div className="step-number">03</div>

              <div className="step-icon">✏️</div>

              <h3>Edita</h3>

              <p>
                Escribe y organiza tu contenido utilizando títulos, sangrías,
                listas, tablas, imágenes y herramientas de formato.
              </p>
            </div>

            <div className="step-card">
              <div className="step-number">04</div>

              <div className="step-icon">💾</div>

              <h3>Guarda y exporta</h3>

              <p>
                Guarda temporalmente tu documento y exporta el resultado final
                en formato Word o PDF.
              </p>
            </div>
          </div>
        </section>

        {/* =====================================================
           NORMAS APA
           ===================================================== */}

        <section className="apa-info-section">
          <div className="section-heading">
            <span>📚</span>

            <p>GUÍA BÁSICA</p>

            <h2>Conoce las normas APA</h2>

            <div className="heading-line" />

            <article className="section-intro">
              Conceptos esenciales para organizar correctamente un documento
              académico.
            </article>
          </div>

          <div className="apa-info-grid">
            <article className="apa-info-card">
              <div className="info-icon">📐</div>

              <h3>Márgenes</h3>

              <p>
                Los documentos académicos pueden configurarse utilizando
                márgenes uniformes de 2,54 cm en los cuatro lados de la página.
              </p>

              <span className="info-tag">2,54 cm</span>
            </article>

            <article className="apa-info-card">
              <div className="info-icon">🔤</div>

              <h3>Tipografía</h3>

              <p>
                La configuración puede utilizar fuentes académicas como Times
                New Roman, Arial o Calibri, dependiendo de los requisitos
                establecidos.
              </p>

              <span className="info-tag">Fuente legible</span>
            </article>

            <article className="apa-info-card">
              <div className="info-icon">↕️</div>

              <h3>Interlineado</h3>

              <p>
                El interlineado doble es una de las configuraciones utilizadas
                para facilitar la lectura y organización del contenido
                académico.
              </p>

              <span className="info-tag">Doble</span>
            </article>

            <article className="apa-info-card">
              <div className="info-icon">⇥</div>

              <h3>Sangría</h3>

              <p>
                Los párrafos del cuerpo del texto pueden utilizar una sangría de
                primera línea de 1,27 cm.
              </p>

              <span className="info-tag">1,27 cm</span>
            </article>

            <article className="apa-info-card">
              <div className="info-icon">🔢</div>

              <h3>Títulos</h3>

              <p>
                Los encabezados permiten organizar el contenido mediante
                diferentes niveles jerárquicos.
              </p>

              <span className="info-tag">Niveles APA</span>
            </article>

            <article className="apa-info-card">
              <div className="info-icon">📄</div>

              <h3>Numeración</h3>

              <p>
                La herramienta permite activar o desactivar la numeración de las
                páginas del documento.
              </p>

              <span className="info-tag">Páginas</span>
            </article>
          </div>
        </section>

        {/* =====================================================
           ALCANCE
           ===================================================== */}

        <section className="notice-section">
          <div className="notice-card">
            <div className="notice-icon">💡</div>

            <div>
              <h3>Importante</h3>

              <p>
                APA Fácil está diseñada para facilitar la aplicación de
                configuraciones básicas de las normas APA. Se recomienda revisar
                siempre las indicaciones específicas de tu institución educativa
                o docente.
              </p>
            </div>
          </div>
        </section>

        {/* =====================================================
           FOOTER
           ===================================================== */}

        <footer className="footer">
          <div className="footer-logo">APA Fácil</div>

          <p>
            Herramienta web para facilitar la organización de documentos
            académicos.
          </p>

          <span>© 2026 APA Fácil</span>
        </footer>
      </main>
    </div>
  );
}

export default App;
