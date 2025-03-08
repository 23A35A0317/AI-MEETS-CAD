import win32com.client

class SolidWorksAI:
    def __init__(self):
        self.swApp = win32com.client.Dispatch("SldWorks.Application")
        self.swApp.Visible = True
        self.model = None
    
    def open_document(self, file_path):
        """Opens a SolidWorks document."""
        self.model = self.swApp.OpenDoc6(file_path, 1, 0, "", 0, 0)
        if self.model:
            print(f"Opened {file_path}")
        else:
            print("Failed to open document")
    
    def create_sketch(self, plane=1):
        """Creates a new sketch on the specified plane."""
        if self.model:
            selMgr = self.model.SelectionManager
            feature = self.model.FeatureManager
            self.model.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, None, 0)
            sketch = feature.InsertSketch2(True)
            print("Sketch mode activated")
        else:
            print("No active document")
    
    def draw_rectangle(self, x, y, width, height):
        """Draws a rectangle in the active sketch."""
        if self.model:
            sketchMgr = self.model.SketchManager
            sketchMgr.CreateCenterRectangle(x, y, 0, x + width, y + height, 0)
            print("Rectangle created")
        else:
            print("No active sketch")
    
    def extrude(self, depth):
        """Extrudes the current sketch to a given depth."""
        feature = self.model.FeatureManager
        feature.FeatureExtrusion2(True, False, False, 0, 0, depth, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
        print(f"Extruded by {depth} units")
    
    def save_as(self, save_path):
        """Saves the current document."""
        self.model.SaveAs(save_path)
        print(f"Saved as {save_path}")
    
    def close_document(self):
        """Closes the current document."""
        if self.model:
            self.swApp.CloseDoc(self.model.GetTitle())
            print("Document closed")
            self.model = None
        else:
            print("No active document")

