import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np

st.set_page_config(
    page_title="Best Supplier Selector",
    page_icon="🎯",
    layout="wide"
)

st.title("🎯 Best Supplier Selector")
st.markdown("**Automatically select the best supplier based on your business rules**")

st.info("""
**How it works:**
1. Upload your complete comparison sheet (with all supplier prices)
2. Set your business rules (thresholds, priorities)
3. Preview recommendations
4. Download with Column E (Best supplier) filled in
""")

# Step 1: Upload comparison sheet
st.subheader("📋 Step 1: Upload Complete Comparison Sheet")

comparison_file = st.file_uploader(
    "Upload your comparison Excel file",
    type=["xlsx", "xls"],
    help="Must have all supplier quotes already added"
)

if comparison_file:
    try:
        df = pd.read_excel(comparison_file, sheet_name='Comparison')
        st.success(f"✅ Loaded {len(df)} parts")
        
        # Show what's available
        with st.expander("📊 Data Preview"):
            preview_cols = ['Part Number', 'Description']
            if 'JB Unit Price' in df.columns:
                preview_cols.append('JB Unit Price')
            if 'Porsche ZA Unit Price' in df.columns:
                preview_cols.append('Porsche ZA Unit Price')
            
            st.dataframe(df[preview_cols].head(10), use_container_width=True)
            
            # Show which suppliers have data
            st.write("**Suppliers with prices:**")
            suppliers_found = []
            if 'JB Unit Price' in df.columns and df['JB Unit Price'].notna().any():
                suppliers_found.append(f"✅ JB ({df['JB Unit Price'].notna().sum()} parts)")
            if 'Porsche ZA Unit Price' in df.columns and df['Porsche ZA Unit Price'].notna().any():
                suppliers_found.append(f"✅ Porsche ZA ({df['Porsche ZA Unit Price'].notna().sum()} parts)")
            
            # Check international suppliers
            for supplier, col_pattern in [
                ("EBS", "EBS Unit Price\n(ZAR+Shipping)"),
                ("PartsWise OEM", "PW OE Unit Price\n(ZAR+Shipping)"),
                ("D911", "D911 Unit Price\n(ZAR+Shipping)")
            ]:
                if col_pattern in df.columns and df[col_pattern].notna().any():
                    suppliers_found.append(f"✅ {supplier} ({df[col_pattern].notna().sum()} parts)")
            
            for s in suppliers_found:
                st.write(s)
        
        st.divider()
        
        # Step 2: Business Rules
        st.subheader("⚙️ Step 2: Set Your Business Rules")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**JB Priority:**")
            jb_threshold = st.slider(
                "JB stays best unless others save more than:",
                min_value=0,
                max_value=2000,
                value=700,
                step=50,
                help="JB is prioritized unless another supplier saves this much"
            )
            st.caption(f"💡 JB chosen unless savings > R{jb_threshold}")
        
        with col2:
            st.write("**International Threshold:**")
            intl_threshold = st.slider(
                "International suppliers must save (%):",
                min_value=0,
                max_value=30,
                value=15,
                step=1,
                help="PartsWise, EBS, D911 must save this % vs PZA (logistics cost)"
            )
            st.caption(f"💡 International must save >{intl_threshold}% vs PZA")
        
        st.write("**Additional Settings:**")
        show_reason = st.checkbox(
            "Show selection reason",
            value=True,
            help="Add explanation of why each supplier was chosen"
        )
        
        st.divider()
        
        # Step 3: Calculate Best Supplier
        st.subheader("🔄 Step 3: Calculating Best Supplier...")
        
        # Define columns
        JB_COL = 'JB Unit Price'
        PZA_COL = 'Porsche ZA Unit Price'
        EBS_SHIP_COL = 'EBS Unit Price\n(ZAR+Shipping)'
        PW_SHIP_COL = 'PW OE Unit Price\n(ZAR+Shipping)'
        D911_SHIP_COL = 'D911 Unit Price\n(ZAR+Shipping)'
        EBS_GEN_COL = 'EBS Genuine'
        
        results = []
        
        for idx, row in df.iterrows():
            # Collect valid prices
            prices = {}
            international_suppliers = []
            
            # JB (local)
            if pd.notna(row.get(JB_COL)) and row.get(JB_COL, 0) > 0:
                prices['JB'] = row[JB_COL]
            
            # Porsche ZA (local baseline)
            pza_price = None
            if pd.notna(row.get(PZA_COL)) and row.get(PZA_COL, 0) > 0:
                prices['Porsche ZA'] = row[PZA_COL]
                pza_price = row[PZA_COL]
            
            # D911 (international)
            if pd.notna(row.get(D911_SHIP_COL)) and row.get(D911_SHIP_COL, 0) > 0:
                prices['D911'] = row[D911_SHIP_COL]
                international_suppliers.append('D911')
            
            # EBS (international) - include all EBS parts
            if pd.notna(row.get(EBS_SHIP_COL)) and row.get(EBS_SHIP_COL, 0) > 0:
                prices['EBS'] = row[EBS_SHIP_COL]
                international_suppliers.append('EBS')
            
            # PartsWise OEM (international)
            if pd.notna(row.get(PW_SHIP_COL)) and row.get(PW_SHIP_COL, 0) > 0:
                prices['PW OEM'] = row[PW_SHIP_COL]
                international_suppliers.append('PW OEM')
            
            # Apply international threshold
            if pza_price:
                threshold_price = pza_price * (1 - intl_threshold / 100)
                
                # Remove international suppliers that don't meet threshold
                for supplier in international_suppliers[:]:  # Copy list
                    if prices[supplier] >= threshold_price:
                        del prices[supplier]
                        international_suppliers.remove(supplier)
            
            # Find cheapest from remaining
            if prices:
                cheapest_supplier = min(prices, key=prices.get)
                cheapest_price = prices[cheapest_supplier]
                
                # Apply JB priority
                if 'JB' in prices and cheapest_supplier != 'JB':
                    jb_price = prices['JB']
                    savings = jb_price - cheapest_price
                    
                    if savings <= jb_threshold:
                        # JB priority applies
                        best_supplier = 'JB'
                        best_price = jb_price
                        reason = f'JB Priority (savings only R{savings:.0f})'
                    else:
                        # Savings worth switching
                        best_supplier = cheapest_supplier
                        best_price = cheapest_price
                        if cheapest_supplier in ['D911', 'EBS', 'PW OEM']:
                            pct_saving = ((pza_price - cheapest_price) / pza_price * 100) if pza_price else 0
                            reason = f'Saves {pct_saving:.1f}% vs PZA (>{intl_threshold}% threshold)'
                        else:
                            reason = f'Saves R{savings:.0f} vs JB (>R{jb_threshold})'
                else:
                    best_supplier = cheapest_supplier
                    best_price = cheapest_price
                    
                    if best_supplier in ['D911', 'EBS', 'PW OEM'] and pza_price:
                        pct_saving = ((pza_price - best_price) / pza_price * 100)
                        reason = f'Saves {pct_saving:.1f}% vs PZA'
                    elif best_supplier == 'JB':
                        reason = 'JB available (preferred)'
                    else:
                        reason = 'Best available price'
                
                # Calculate savings vs PZA
                if pza_price:
                    savings_vs_pza = pza_price - best_price
                    pct_vs_pza = (savings_vs_pza / pza_price * 100)
                else:
                    savings_vs_pza = None
                    pct_vs_pza = None
            else:
                best_supplier = None
                best_price = None
                reason = 'No qualifying suppliers'
                savings_vs_pza = None
                pct_vs_pza = None
            
            results.append({
                'Best supplier': best_supplier,
                'Best Price': best_price,
                'Reason': reason,
                'Savings vs PZA': savings_vs_pza,
                '% Savings': pct_vs_pza
            })
        
        # Add results to dataframe
        results_df = pd.DataFrame(results)
        df['Best supplier'] = results_df['Best supplier']
        df['Best Price (Calculated)'] = results_df['Best Price']
        
        if show_reason:
            df['Selection Reason'] = results_df['Reason']
        
        df['Savings vs PZA (R)'] = results_df['Savings vs PZA']
        df['Savings vs PZA (%)'] = results_df['% Savings']
        
        st.success("✅ Best supplier calculated for all parts!")
        
        # Summary
        st.divider()
        st.subheader("📊 Step 4: Results Summary")
        
        supplier_counts = df['Best supplier'].value_counts()
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total Parts", len(df))
        
        with col2:
            parts_with_supplier = df['Best supplier'].notna().sum()
            st.metric("Parts with Supplier", parts_with_supplier)
        
        with col3:
            total_cost = (df['Best Price (Calculated)'] * df['Quantity']).sum()
            st.metric("Total Cost", f"R {total_cost:,.0f}" if pd.notna(total_cost) else "N/A")
        
            if pd.notna(supplier):
                count = supplier_counts[supplier]
                supplier_parts = df[df['Best supplier'] == supplier]
                total_value = (supplier_parts['Best Price (Calculated)'] * supplier_parts['Quantity']).sum()
                pct = (count / parts_with_supplier * 100) if parts_with_supplier > 0 else 0
                
                breakdown.append({
                    'Supplier': supplier,
                    'Parts': count,
                    'Percentage': f"{pct:.1f}%",
                    'Total Value': f"R {total_value:,.0f}"
                })
        
        breakdown_df = pd.DataFrame(breakdown)
        st.dataframe(breakdown_df, use_container_width=True, hide_index=True)
        
        # Detailed view
        st.divider()
        st.subheader("📋 Detailed Results")
        
        col1, col2 = st.columns(2)
        
        with col1:
            search = st.text_input("🔍 Search part number or description")
        
        with col2:
            filter_supplier = st.selectbox(
                "Filter by supplier",
                options=["All"] + list(df['Best supplier'].dropna().unique())
            )
        
        # Apply filters
        display_df = df.copy()
        
        if search:
            mask = (
                display_df['Part Number'].astype(str).str.contains(search, case=False, na=False) |
                display_df['Description'].astype(str).str.contains(search, case=False, na=False)
            )
            display_df = display_df[mask]
        
        if filter_supplier != "All":
            display_df = display_df[display_df['Best supplier'] == filter_supplier]
        
        # Select display columns
        display_cols = ['Part Number', 'Description', 'Quantity']
        
        if JB_COL in df.columns:
            display_cols.append(JB_COL)
        if PZA_COL in df.columns:
            display_cols.append(PZA_COL)
        
        display_cols.extend(['Best supplier', 'Best Price (Calculated)'])
        
        if show_reason:
            display_cols.append('Selection Reason')
        
        display_cols.extend(['Savings vs PZA (R)', 'Savings vs PZA (%)'])
        
        # Format for display
        formatted_df = display_df[display_cols].copy()
        
        for col in formatted_df.columns:
            if 'Price' in col or 'Savings' in col and '(%)' not in col:
                formatted_df[col] = formatted_df[col].apply(
                    lambda x: f"R {x:,.2f}" if pd.notna(x) and x != 0 else "-" if pd.isna(x) else "R 0"
                )
            elif '(%)' in col:
                formatted_df[col] = formatted_df[col].apply(
                    lambda x: f"{x:.1f}%" if pd.notna(x) else "-"
                )
        
        st.dataframe(formatted_df, use_container_width=True, height=500)
        st.caption(f"Showing {len(display_df)} of {len(df)} parts")
        
        # Export
        st.divider()
        st.subheader("📥 Step 5: Download with Best Supplier")
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.write("**Download your comparison sheet with Column E (Best supplier) filled in.**")
            st.info(f"""
            ✅ Column E updated with best supplier
            ✅ Business rules applied (JB priority: R{jb_threshold}, International: {intl_threshold}%)
            ✅ All EBS parts included (Genuine & Aftermarket)
            ✅ Row order maintained
            ✅ Ready to copy/paste or use directly
            """)
        
        with col2:
            # Create export
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Comparison', index=False)
            
            output.seek(0)
            
            st.download_button(
                label="📥 Download",
                data=output,
                file_name="comparison_with_best_supplier.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
        import traceback
        st.error(traceback.format_exc())

else:
    st.info("👈 Upload your complete comparison sheet to get started")
    
    st.markdown("""
    ---
    ### How to use this app:
    
    **Prerequisites:**
    1. Complete comparison sheet with all supplier quotes added
    2. Use the Quote Matcher app first to add all suppliers
    
    **This app will:**
    1. Apply your business rules automatically
    2. Calculate best supplier for each part
    3. Fill Column E (Best supplier)
    4. Show cost savings
    
    **Your Business Rules:**
    - ✅ **JB Priority:** Choose JB unless others save >R700
    - ✅ **International Threshold:** Must save >15% vs PZA (logistics)
    - ✅ **All EBS Parts:** Includes both Genuine & Aftermarket
    - ✅ **Local Preference:** Prefer JB/PZA when close
    
    **Adjustable:**
    - Change thresholds with sliders
    - Test different scenarios
    - Compare before downloading
    """)
