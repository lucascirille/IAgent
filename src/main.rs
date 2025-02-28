use anyhow::{Context, Result};
use calamine::{open_workbook, Reader, Xlsx};
use dotenv::dotenv;
use reqwest::Client;
use rust_xlsxwriter::{Workbook};
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::collections::HashMap;
use std::env;
use std::io::{self, BufRead, Write};
use std::path::Path;
use tokio;

// Estructuras para la API de Deepseek
#[derive(Serialize, Debug)]
struct Message {
    role: String,
    content: String,
}

#[derive(Deserialize, Debug)]
struct DeepseekResponse {
    choices: Vec<DeepseekChoice>,
}

#[derive(Deserialize, Debug)]
struct DeepseekChoice {
    message: DeepseekMessage,
}

#[derive(Deserialize, Debug)]
struct DeepseekMessage {
    content: String,
}

// Enum para comandos de Excel
enum ExcelCommand {
    ReadFile(String),
    CreateFile(String),
    WriteData(String, String),
}

#[tokio::main]
async fn main() -> Result<()> {
    // Cargar variables de entorno desde un archivo .env
    dotenv().ok();
    let api_key = env::var("DEEPSEEK_API_KEY")
        .context("No se encontró DEEPSEEK_API_KEY en el entorno")?;
    let api_url = env::var("DEEPSEEK_API_URL")
        .unwrap_or_else(|_| "https://api.deepseek.com/v1/chat/completions".to_string());

    println!("=== Agente de IA con Deepseek para Excel ===");
    println!("Escribe 'ayuda' para ver comandos disponibles");
    println!("Escribe 'salir' para terminar");

    // Historial de conversaciones para el contexto
    let mut conversation_history: Vec<Message> = vec![Message {
        role: "system".to_string(),
        content: "Eres un asistente especializado en manipular archivos Excel. Puedes analizar datos, crear gráficos, realizar cálculos y generar informes basados en datos de Excel. Responde de manera concisa y enfocada en la tarea solicitada.".to_string(),
    }];

    let client = Client::new();
    let stdin = io::stdin();
    let mut reader = stdin.lock();

    loop {
        print!("> ");
        io::stdout().flush()?;
        
        let mut input = String::new();
        reader.read_line(&mut input)?;
        let input = input.trim();

        if input.eq_ignore_ascii_case("salir") {
            println!("Adiós!");
            break;
        }

        if input.eq_ignore_ascii_case("ayuda") {
            show_help();
            continue;
        }

        // Detecta si hay comandos específicos para Excel
        if let Some(command) = parse_excel_command(input) {
            match command {
                ExcelCommand::ReadFile(filename) => {
                    match read_excel_file(&filename) {
                        Ok(data) => {
                            println!("✅ Archivo leído correctamente");
                            // Convertimos los datos a un formato más amigable para el contexto
                            let data_summary = summarize_excel_data(&data);
                            conversation_history.push(Message {
                                role: "system".to_string(),
                                content: format!("Datos del archivo Excel '{}': {}", filename, data_summary),
                            });
                        }
                        Err(e) => println!("❌ Error al leer el archivo: {}", e),
                    }
                }
                ExcelCommand::CreateFile(filename) => {
                    match create_excel_file(&filename) {
                        Ok(_) => println!("✅ Archivo creado correctamente: {}", filename),
                        Err(e) => println!("❌ Error al crear el archivo: {}", e),
                    }
                }
                ExcelCommand::WriteData(filename, data) => {
                    match write_excel_data(&filename, &data) {
                        Ok(_) => println!("✅ Datos escritos correctamente en {}", filename),
                        Err(e) => println!("❌ Error al escribir datos: {}", e),
                    }
                }
            }
            continue;
        }

        // Añade la entrada del usuario al historial
        conversation_history.push(Message {
            role: "user".to_string(),
            content: input.to_string(),
        });

        // Obtiene respuesta de Deepseek
        match get_deepseek_response(&client, &api_url, &api_key, &conversation_history).await {
            Ok(response) => {
                println!("{}", response);
                // Añade la respuesta al historial
                conversation_history.push(Message {
                    role: "assistant".to_string(),
                    content: response,
                });
            }
            Err(e) => println!("Error al comunicarse con Deepseek: {}", e),
        }
    }

    Ok(())
}

// Función para mostrar ayuda
fn show_help() {
    println!("Comandos disponibles:");
    println!("  leer_excel <archivo.xlsx> - Lee un archivo Excel");
    println!("  crear_excel <archivo.xlsx> - Crea un nuevo archivo Excel");
    println!("  escribir_excel <archivo.xlsx> <datos> - Escribe datos en un archivo Excel");
    println!("  ayuda - Muestra esta información");
    println!("  salir - Termina el programa");
    println!();
    println!("También puedes hacer preguntas sobre manipulación de Excel o solicitar ayuda.");
}

// Parsea comandos específicos de Excel
fn parse_excel_command(input: &str) -> Option<ExcelCommand> {
    let parts: Vec<&str> = input.split_whitespace().collect();
    
    match parts.get(0) {
        Some(&"leer_excel") if parts.len() >= 2 => {
            Some(ExcelCommand::ReadFile(parts[1].to_string()))
        }
        Some(&"crear_excel") if parts.len() >= 2 => {
            Some(ExcelCommand::CreateFile(parts[1].to_string()))
        }
        Some(&"escribir_excel") if parts.len() >= 3 => {
            let filename = parts[1].to_string();
            let data = parts[2..].join(" ");
            Some(ExcelCommand::WriteData(filename, data))
        }
        _ => None,
    }
}

// Función para leer un archivo Excel
fn read_excel_file(filename: &str) -> Result<HashMap<String, Vec<Vec<String>>>> {
    let path = Path::new(filename);
    let mut workbook: Xlsx<_> = open_workbook(path)
        .context(format!("No se pudo abrir el archivo {}", filename))?;
    let mut result = HashMap::new();

    for sheet_name in workbook.sheet_names().to_owned() {
        if let Some(Ok(range)) = workbook.worksheet_range(&sheet_name) {
            let mut sheet_data = Vec::new();
            for row in range.rows() {
                let row_data: Vec<String> = row
                    .iter()
                    .map(|cell| cell.to_string())
                    .collect();
                sheet_data.push(row_data);
            }
            result.insert(sheet_name, sheet_data);
        }
    }

    Ok(result)
}

// Función para crear un resumen simplificado de los datos de Excel
fn summarize_excel_data(data: &HashMap<String, Vec<Vec<String>>>) -> String {
    let mut summary = String::new();
    
    for (sheet_name, rows) in data {
        summary.push_str(&format!("Hoja: {} ({} filas)\n", sheet_name, rows.len()));
        
        // Añadir encabezados si existen
        if !rows.is_empty() {
            summary.push_str("Encabezados: ");
            summary.push_str(&rows[0].join(", "));
            summary.push_str("\n");
        }
        
        // Limitar a mostrar solo algunas filas para no sobrecargar el contexto
        let max_rows = std::cmp::min(5, rows.len());
        if max_rows > 1 {
            summary.push_str("Primeras filas de datos:\n");
            for i in 1..max_rows {
                summary.push_str(&format!("  {}\n", rows[i].join(", ")));
            }
        }
    }
    
    summary
}

// Función para crear un archivo Excel
fn create_excel_file(filename: &str) -> Result<()> {
    let mut workbook = Workbook::new();
    let _worksheet = workbook.add_worksheet();
    
    workbook.save(filename)?;
    Ok(())
}

// Función para escribir datos en un archivo Excel
fn write_excel_data(filename: &str, data: &str) -> Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Parseamos los datos (formato simple: filas separadas por punto y coma, columnas por coma)
    for (row_idx, line) in data.split(';').enumerate() {
        for (col_idx, value) in line.split(',').enumerate() {
            worksheet.write_string(row_idx as u32, col_idx as u16, value.trim())?;
        }
    }

    workbook.save(filename)?;
    Ok(())
}

// Función para obtener una respuesta de Deepseek
async fn get_deepseek_response(
    client: &Client,
    api_url: &str,
    api_key: &str,
    messages: &[Message],
) -> Result<String> {
    let request_body = json!({
        "model": "deepseek-coder", // Ajusta según el modelo disponible
        "messages": messages,
        "temperature": 0.7,
        "max_tokens": 500
    });

    let response = client
        .post(api_url)
        .header("Authorization", format!("Bearer {}", api_key))
        .header("Content-Type", "application/json")
        .json(&request_body)
        .send()
        .await?;

    if response.status().is_success() {
        let response_data: DeepseekResponse = response.json().await?;
        if let Some(choice) = response_data.choices.get(0) {
            return Ok(choice.message.content.clone());
        }
    }

    Err(anyhow::anyhow!("No se pudo obtener una respuesta válida de Deepseek"))
}
