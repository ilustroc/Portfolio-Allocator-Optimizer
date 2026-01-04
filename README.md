# Optimizador de Asignación de Cartera (Portfolio Allocator)

Este proyecto automatiza la distribución estratégica de clientes (deudores) a diferentes agencias de cobranza (gestores externos). Utiliza Programación Lineal Entera (CP-SAT) mediante Google OR-Tools para asegurar equidad y cumplimiento de reglas de negocio.

## 🚀 Características

- **Distribución Equitativa:** Balancea la carga entre empresas basándose en el **Capital (Deuda)** y el **Número de Cuentas**.
- **Regla de Retención (Veto):** Limita la cantidad de clientes que pueden permanecer con su gestor anterior (Máximo 20%), forzando una rotación del 80% de la cartera.
- **Optimización Geográfica:** Intenta mantener equidad en la distribución por zonas (Norte, Sur, Lima, etc.).
- **Manejo de "Inalterables":** Respeta asignaciones fijas marcadas previamente en la base de datos.

## 📋 Requisitos

- Python 3.8+
- Librerías listadas en `requirements.txt`

## 🛠️ Instalación

1. Clona el repositorio:
   ```bash
   git clone [https://github.com/ilustroc/portfolio-allocator.git](https://github.com/TU_USUARIO/portfolio-allocator.git)