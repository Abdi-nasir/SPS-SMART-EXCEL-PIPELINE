#!/bin/bash
# push-to-docker-hub.sh - Push only SPS Excel Dashboard image

# Configuration
DOCKER_USER="tamhiid"  # Replace with your actual Docker Hub username
PROJECT_NAME="sps-board-dashboard"
IMAGE_NAME="sps-excel-dashboard"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Function to print colored output
print_message() {
    echo -e "${2}${1}${NC}"
}

# Function to validate version format
validate_version() {
    if [[ $1 =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
        return 0
    else
        return 1
    fi
}

# Check if docker is running
if ! docker info >/dev/null 2>&1; then
    print_message "Docker is not running. Please start Docker first." "$RED"
    exit 1
fi

print_message " Starting push process for $IMAGE_NAME..." "$GREEN"

# Ask for version number
print_message "\n Enter version number (format: MAJOR.MINOR.PATCH e.g., 1.0.0):" "$YELLOW"
read -p "Version: " VERSION

# Validate version format
if ! validate_version "$VERSION"; then
    print_message "❌ Invalid version format. Please use MAJOR.MINOR.PATCH (e.g., 1.0.0)" "$RED"
    exit 1
fi

print_message "✅ Using version: $VERSION" "$GREEN"

# Show version suggestion based on last version
if [ -f ".last-version" ]; then
    LAST_VERSION=$(cat .last-version)
    print_message "\n Last version was: $LAST_VERSION" "$YELLOW"
    print_message "   Suggested next version:" "$YELLOW"
    
    # Parse last version
    IFS='.' read -r MAJOR MINOR PATCH <<< "$LAST_VERSION"
    echo "   - Patch update (bug fix): $MAJOR.$MINOR.$((PATCH + 1))"
    echo "   - Minor update (new feature): $MAJOR.$((MINOR + 1)).0"
    echo "   - Major update (breaking change): $((MAJOR + 1)).0.0"
fi

# Confirm version
read -p "Continue with version $VERSION? (y/n): " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    print_message "❌ Aborted by user" "$RED"
    exit 1
fi

# Step 1: Build the dashboard image only
print_message "\n Building $IMAGE_NAME image..." "$YELLOW"
docker build -t ${PROJECT_NAME}_${IMAGE_NAME}:latest .

if [ $? -ne 0 ]; then
    print_message "❌ Build failed. Aborting." "$RED"
    exit 1
fi
print_message "✅ Build completed successfully!" "$GREEN"

# Step 2: Login to Docker Hub
print_message "\n Logging in to Docker Hub..." "$YELLOW"
docker login

if [ $? -ne 0 ]; then
    print_message "❌ Login failed. Aborting." "$RED"
    exit 1
fi
print_message "✅ Login successful!" "$GREEN"

# Step 3: Tag and push the dashboard image
print_message "\n Tagging and pushing $IMAGE_NAME..." "$YELLOW"

# Try different possible local image names
LOCAL_IMAGE_CANDIDATES=(
    "${PROJECT_NAME}_${IMAGE_NAME}:latest"
    "${IMAGE_NAME}:latest"
    "${PROJECT_NAME}-${IMAGE_NAME}:latest"
)

LOCAL_IMAGE=""
for CANDIDATE in "${LOCAL_IMAGE_CANDIDATES[@]}"; do
    if docker image inspect $CANDIDATE >/dev/null 2>&1; then
        LOCAL_IMAGE=$CANDIDATE
        print_message "   Found local image: $LOCAL_IMAGE" "$GREEN"
        break
    fi
done

if [ -z "$LOCAL_IMAGE" ]; then
    print_message "❌ Could not find local image for $IMAGE_NAME. Available images:" "$RED"
    docker images | head -10
    exit 1
fi

# Tag with version and latest
REMOTE_IMAGE_VERSION="${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}:${VERSION}"
REMOTE_IMAGE_LATEST="${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}:latest"

print_message "   Tagging: $LOCAL_IMAGE -> $REMOTE_IMAGE_VERSION" 
docker tag $LOCAL_IMAGE $REMOTE_IMAGE_VERSION

print_message "   Tagging: $LOCAL_IMAGE -> $REMOTE_IMAGE_LATEST"
docker tag $LOCAL_IMAGE $REMOTE_IMAGE_LATEST

if [ $? -ne 0 ]; then
    print_message "❌ Failed to tag $IMAGE_NAME" "$RED"
    exit 1
fi

# Push both tags
print_message "   Pushing version: $REMOTE_IMAGE_VERSION"
docker push $REMOTE_IMAGE_VERSION

print_message "   Pushing latest: $REMOTE_IMAGE_LATEST"
docker push $REMOTE_IMAGE_LATEST

if [ $? -eq 0 ]; then
    print_message "✅ $IMAGE_NAME pushed successfully! (version: $VERSION)" "$GREEN"
else
    print_message "❌ Failed to push $IMAGE_NAME" "$RED"
    exit 1
fi

# Step 4: Generate docker-compose file for the dashboard only
print_message "\n Generating docker-compose file..." "$YELLOW"

cat > docker-compose.prod.yml << EOF
# Production docker-compose file for SPS Excel Dashboard
# Generated on $(date)
# Version: $VERSION
# This file uses images from Docker Hub

version: '3.8'

services:
  sps-dashboard:
    image: ${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}:${VERSION}
    container_name: sps-excel-dashboard
    ports:
      - "8501:8501"
    volumes:
      # Persist uploaded files and processed data
      - ./data:/app/data
      - ./reports:/app/reports
    environment:
      - STREAMLIT_SERVER_MAX_UPLOAD_SIZE=200
      - STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
      - TZ=UTC
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:8501/_stcore/health"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 40s
    logging:
      driver: "json-file"
      options:
        max-size: "10m"
        max-file: "3"

# Optional: Add nginx reverse proxy (uncomment if needed)
#  nginx:
#    image: nginx:alpine
#    container_name: sps-nginx
#    ports:
#      - "80:80"
#      - "443:443"
#    volumes:
#      - ./nginx.conf:/etc/nginx/nginx.conf:ro
#      - ./ssl:/etc/nginx/ssl:ro
#    depends_on:
#      - sps-dashboard
#    restart: unless-stopped
EOF

print_message " Generated docker-compose.prod.yml" "$GREEN"

# Step 5: Generate simple docker run command script
cat > deploy.sh << EOF
#!/bin/bash
# Deploy script for SPS Excel Dashboard
# Version: $VERSION

# Pull the latest image
docker pull ${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}:${VERSION}

# Stop and remove existing container if running
docker stop sps-excel-dashboard 2>/dev/null
docker rm sps-excel-dashboard 2>/dev/null

# Create necessary directories
mkdir -p data/raw data/processed reports

# Run the container
docker run -d \\
  --name sps-excel-dashboard \\
  -p 8501:8501 \\
  -v \$(pwd)/data:/app/data \\
  -v \$(pwd)/reports:/app/reports \\
  --restart unless-stopped \\
  ${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}:${VERSION}

echo " SPS Excel Dashboard deployed successfully!"
echo " Access it at: http://localhost:8501"
EOF

chmod +x deploy.sh
print_message " Generated deploy.sh script" "$GREEN"

# Step 6: Save version for next time
echo "$VERSION" > .last-version
print_message " Saved version $VERSION to .last-version" "$GREEN"

# Step 7: Create a version changelog
echo "$(date): Released version $VERSION - SPS Excel Dashboard" >> CHANGELOG.txt
print_message " Updated CHANGELOG.txt" "$GREEN"

# Step 8: Summary
print_message "\n Process completed successfully!" "$GREEN"
print_message "\n Pushed images (version: $VERSION):" "$YELLOW"
echo "   - ${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}:${VERSION}"
echo "   - ${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}:latest"

print_message "\n Next steps:" "$YELLOW"
print_message "   1. Update CHANGELOG.txt with what's new in version $VERSION" "$NC"
print_message "   2. Test the image locally:" "$NC"
print_message "      docker run -d -p 8501:8501 ${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}:${VERSION}" "$NC"
print_message "   3. Deploy to server:" "$NC"
print_message "      scp docker-compose.prod.yml deploy.sh your_user@your_server:~/sps-dashboard/" "$NC"
print_message "      ssh your_user@your_server" "$NC"
print_message "      cd ~/sps-dashboard && ./deploy.sh" "$NC"

print_message "\n📌 Current version: $VERSION" "$GREEN"
print_message "\n🔗 Image URL: https://hub.docker.com/r/${DOCKER_USER}/${PROJECT_NAME}-${IMAGE_NAME}" "$BLUE"