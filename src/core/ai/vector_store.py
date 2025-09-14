"""
Vector store for embeddings and semantic search of presentation content.
"""

import os
import json
from typing import List, Dict, Optional, Tuple
from pathlib import Path
import numpy as np


class VectorStore:
    """Local vector database for presentation content embeddings."""
    
    def __init__(self, storage_path: Optional[str] = None):
        self.storage_path = Path(storage_path) if storage_path else Path("assets/models/vector_store")
        self.storage_path.mkdir(parents=True, exist_ok=True)
        
        self.embeddings_file = self.storage_path / "embeddings.json"
        self.metadata_file = self.storage_path / "metadata.json"
        
        self.embeddings: Dict[str, List[float]] = {}
        self.metadata: Dict[str, Dict] = {}
        
        self._load_data()
    
    def _load_data(self):
        """Load embeddings and metadata from disk."""
        try:
            if self.embeddings_file.exists():
                with open(self.embeddings_file, 'r') as f:
                    self.embeddings = json.load(f)
            
            if self.metadata_file.exists():
                with open(self.metadata_file, 'r') as f:
                    self.metadata = json.load(f)
            
            print(f"Loaded vector store with {len(self.embeddings)} embeddings")
        except Exception as e:
            print(f"Failed to load vector store: {e}")
            self.embeddings = {}
            self.metadata = {}
    
    def _save_data(self):
        """Save embeddings and metadata to disk."""
        try:
            with open(self.embeddings_file, 'w') as f:
                json.dump(self.embeddings, f, indent=2)
            
            with open(self.metadata_file, 'w') as f:
                json.dump(self.metadata, f, indent=2)
        except Exception as e:
            print(f"Failed to save vector store: {e}")
    
    def add_document(self, doc_id: str, text: str, metadata: Optional[Dict] = None):
        """Add a document to the vector store."""
        try:
            # Generate embedding (placeholder implementation)
            embedding = self._generate_embedding(text)
            
            self.embeddings[doc_id] = embedding
            self.metadata[doc_id] = {
                'text': text,
                'timestamp': time.time(),
                **(metadata or {})
            }
            
            self._save_data()
            print(f"📝 Added document {doc_id} to vector store")
        except Exception as e:
            print(f"Failed to add document: {e}")
    
    def _generate_embedding(self, text: str) -> List[float]:
        """Generate embedding for text (placeholder implementation)."""
        # TODO: Implement actual embedding generation
        # This could use sentence-transformers, OpenAI embeddings, or other models
        
        # Placeholder: simple hash-based "embedding"
        import hashlib
        hash_obj = hashlib.md5(text.encode())
        hash_bytes = hash_obj.digest()
        
        # Convert to float vector
        embedding = []
        for i in range(0, len(hash_bytes), 4):
            chunk = hash_bytes[i:i+4]
            if len(chunk) == 4:
                val = int.from_bytes(chunk, 'big') / (2**32)
                embedding.append(val)
        
        # Pad or truncate to fixed size
        target_size = 384  # Common embedding size
        while len(embedding) < target_size:
            embedding.append(0.0)
        embedding = embedding[:target_size]
        
        return embedding
    
    def search(self, query: str, top_k: int = 5) -> List[Tuple[str, float, Dict]]:
        """Search for similar documents."""
        try:
            query_embedding = self._generate_embedding(query)
            results = []
            
            for doc_id, doc_embedding in self.embeddings.items():
                similarity = self._cosine_similarity(query_embedding, doc_embedding)
                results.append((doc_id, similarity, self.metadata.get(doc_id, {})))
            
            # Sort by similarity and return top_k
            results.sort(key=lambda x: x[1], reverse=True)
            return results[:top_k]
        except Exception as e:
            print(f"Search failed: {e}")
            return []
    
    def _cosine_similarity(self, vec1: List[float], vec2: List[float]) -> float:
        """Calculate cosine similarity between two vectors."""
        try:
            vec1 = np.array(vec1)
            vec2 = np.array(vec2)
            
            dot_product = np.dot(vec1, vec2)
            norm1 = np.linalg.norm(vec1)
            norm2 = np.linalg.norm(vec2)
            
            if norm1 == 0 or norm2 == 0:
                return 0.0
            
            return dot_product / (norm1 * norm2)
        except Exception:
            return 0.0
    
    def add_presentation_slides(self, presentation_id: str, slides_data: List[Dict]):
        """Add all slides from a presentation to the vector store."""
        for i, slide_data in enumerate(slides_data):
            slide_id = f"{presentation_id}_slide_{i+1}"
            text_content = slide_data.get('text_content', '')
            
            if text_content.strip():
                metadata = {
                    'presentation_id': presentation_id,
                    'slide_number': i + 1,
                    'slide_type': slide_data.get('type', 'content'),
                    'ocr_text': slide_data.get('ocr_text', ''),
                    'object_count': slide_data.get('object_count', 0)
                }
                self.add_document(slide_id, text_content, metadata)
    
    def get_slide_context(self, presentation_id: str, current_slide: int, context_window: int = 2) -> str:
        """Get contextual information around the current slide."""
        context_slides = []
        
        for i in range(max(1, current_slide - context_window), 
                      min(current_slide + context_window + 1, 100)):  # Assume max 100 slides
            slide_id = f"{presentation_id}_slide_{i}"
            if slide_id in self.metadata:
                slide_text = self.metadata[slide_id].get('text', '')
                if slide_text.strip():
                    context_slides.append(f"Slide {i}: {slide_text}")
        
        return '\n\n'.join(context_slides)
    
    def clear_presentation(self, presentation_id: str):
        """Remove all slides for a specific presentation."""
        slides_to_remove = []
        for doc_id in self.embeddings.keys():
            if doc_id.startswith(f"{presentation_id}_slide_"):
                slides_to_remove.append(doc_id)
        
        for doc_id in slides_to_remove:
            self.embeddings.pop(doc_id, None)
            self.metadata.pop(doc_id, None)
        
        if slides_to_remove:
            self._save_data()
            print(f"Removed {len(slides_to_remove)} slides for presentation {presentation_id}")


# Import time for metadata timestamp
import time
