//
//  ContentView.swift
//  SearchInExcel
//
//  Created by Moy on 22/06/21.
//

import SwiftUI
import CoreXLSX

struct ContentView: View {
    
    @State var folder_selected: String = ""
    @State var search_string: String = ""
    @State var log: String = "Results"
    
    var body: some View {
        
        VStack{
            HStack{
                
                TextField("/home", text: $folder_selected).frame(maxWidth: .infinity, maxHeight: 30)
                
                Button(action:{
                    //
                    selectFolder()
                }){
                    Text("Select Folder")
                }.frame(minWidth: 100)
            }
            
            HStack{
                
                TextField("Search", text: $search_string).frame(maxWidth: .infinity, maxHeight: 30)
                
                Button(action:{
                    search()
                }){
                    Text("Search")
                }.frame(minWidth: 100)
            }
            
            if #available(macOS 11.0, *) {
                TextEditor(text: $log).frame(maxWidth: .infinity, maxHeight: .infinity)
            } else {
                // Fallback on earlier versions
            }
            
        }.frame(maxWidth: .infinity, maxHeight: .infinity)
    }
    
    func selectFolder(){
        let dialog = NSOpenPanel();

        dialog.title                   = "Choose a folder";
        dialog.showsResizeIndicator    = true;
        dialog.showsHiddenFiles        = false;
        dialog.allowsMultipleSelection = false;
        dialog.canChooseFiles         = false;
        dialog.canChooseDirectories = true;

        if (dialog.runModal() ==  NSApplication.ModalResponse.OK) {
            let result = dialog.url // Pathname of the file

            if (result != nil) {
                let path: String = result!.path
                
                folder_selected = path
                
            }
            
        } else {
            // User clicked on "Cancel"
            return
        }
    }
    
    func search(){
        
        log = ""
        
        let fm = FileManager.default
        let path = folder_selected

        do {
            let items = try fm.contentsOfDirectory(atPath: path)

            for item in items {
                if item.contains("xlsx") {
                    print("Found \(item)")
                    searchInFile(filepath: folder_selected+"/"+item)
                }
                
            }
        } catch {
            // failed to read directory – bad permissions, perhaps?
        }
        
    }
    
    func searchInFile(filepath: String){
        
        log += "\n"+filepath
        
        guard let file = XLSXFile(filepath: filepath) else {
            
          log += "\n"+"XLSX file at \(filepath) is corrupted or does not exist"
          fatalError("XLSX file at \(filepath) is corrupted or does not exist")
        }

        for wbk in try! file.parseWorkbooks() {
          for (name, path) in try! file.parseWorksheetPathsAndNames(workbook: wbk) {
            if let worksheetName = name {
              print("This worksheet has a name: \(worksheetName)")
                log += "\n"+"This worksheet has a name: \(worksheetName)"
            }

            let worksheet = try! file.parseWorksheet(at: path)
            for row in worksheet.data?.rows ?? [] {
              for c in row.cells {
                print(c)
                
                var cell_value = c.value!
                var cell_value_lowercase = c.value!.lowercased()
                let search_lowercase = search_string.lowercased()
                
                if(c.type != nil){
                    if(c.type! == CoreXLSX.CellType.sharedString){
                        if let sharedStrings = try! file.parseSharedStrings() {
                            cell_value_lowercase = (c.stringValue(sharedStrings)?.lowercased())!
                            cell_value = (c.stringValue(sharedStrings))!
                        }
                    }
                }
                
                if(cell_value_lowercase.contains(search_lowercase)){
                    log += "\n \(c.reference): \(cell_value)"
                }
                
              }
            }
          }
        }
    }
}


struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}
